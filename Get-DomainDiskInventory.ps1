<#
.SYNOPSIS
    Собирает данные о жестких дисках (модель, серийный и электронный номера) со всех
    компьютеров домена Active Directory и формирует отчеты CSV и HTML.

.DESCRIPTION
    1. Получает список компьютеров из AD (Get-ADComputer), либо из параметра -ComputerName,
       либо из текстового файла -ListPath (по одному имени в строке).
    2. Параллельно опрашивает каждый компьютер через CIM (сначала WSMan, при неудаче - DCOM):
         - Win32_DiskDrive                   : модель, серийный номер, PNPDeviceID, интерфейс, размер
         - Win32_PhysicalMedia               : серийный номер носителя (запасной источник)
         - MSFT_PhysicalDisk (Storage)       : серийный номер, UniqueId (электронный идентификатор),
                                               тип носителя SSD/HDD, состояние, прошивка
    3. Сохраняет результаты:
         DiskInventory_<дата>.csv    - все диски (разделитель ";", UTF-8 с BOM - открывается в Excel)
         DiskInventory_<дата>.html   - сводный отчет
         DiskInventory_<дата>_errors.csv - компьютеры, которые не удалось опросить

    Серийный номер: берется из MSFT_PhysicalDisk, иначе из Win32_DiskDrive, иначе из
    Win32_PhysicalMedia. Если драйвер отдает его в виде HEX-строки, она декодируется.
    Электронный номер: UniqueId (WWN/NAA/EUI/SCSI-идентификатор диска) и PNPDeviceID.

.PARAMETER ComputerName
    Явный список компьютеров. Если не указан - берутся из AD или из -ListPath.

.PARAMETER ListPath
    Текстовый файл со списком компьютеров (по одному в строке).

.PARAMETER SearchBase
    OU в AD, из которой выбирать компьютеры, например "OU=Office,DC=corp,DC=local".

.PARAMETER OperatingSystem
    Маска ОС для фильтрации в AD. По умолчанию "*" (все). Пример: "*Windows 10*", "*Server*".

.PARAMETER IncludeDisabled
    Включать в опрос отключенные учетные записи компьютеров (по умолчанию пропускаются).

.PARAMETER ExcludeUsb
    Не включать в отчет диски с интерфейсом USB.

.PARAMETER Credential
    Учетные данные администратора, если текущая учетная запись не имеет прав на удаленных ПК.

.PARAMETER OutputDir
    Папка для отчетов. По умолчанию - подпапка Reports рядом со скриптом
    (если скрипт запущен не из файла - в текущей папке).

.PARAMETER ThrottleLimit
    Сколько компьютеров опрашивать одновременно. По умолчанию 32.

.PARAMETER TimeoutSec
    Таймаут CIM-операций на один компьютер, сек. По умолчанию 30.

.EXAMPLE
    .\Get-DomainDiskInventory.ps1

.EXAMPLE
    .\Get-DomainDiskInventory.ps1 -SearchBase "OU=Workstations,DC=corp,DC=local" -ExcludeUsb

.EXAMPLE
    .\Get-DomainDiskInventory.ps1 -ComputerName PC01,PC02 -Credential (Get-Credential)

.NOTES
    Требуется: PowerShell 5.1+, модуль RSAT ActiveDirectory (если список берется из AD),
    права локального администратора на опрашиваемых компьютерах, доступ к WinRM (5985)
    или DCOM/RPC (135 + динамические порты), ICMP-ответ от компьютеров.
#>
[CmdletBinding()]
param(
    [string[]]$ComputerName,
    [string]$ListPath,
    [string]$SearchBase,
    [string]$OperatingSystem = '*',
    [switch]$IncludeDisabled,
    [switch]$ExcludeUsb,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$OutputDir,
    [ValidateRange(1, 256)][int]$ThrottleLimit = 32,
    [ValidateRange(5, 600)][int]$TimeoutSec = 30
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# $PSScriptRoot пуст, если код запущен не из файла (вставка в консоль, F8 в ISE, dot-sourcing)
if (-not $OutputDir) {
    $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputDir = Join-Path $baseDir 'Reports'
}

#region Список компьютеров
if ($ComputerName) {
    $computers = $ComputerName
}
elseif ($ListPath) {
    if (-not (Test-Path -LiteralPath $ListPath)) { throw "Файл со списком не найден: $ListPath" }
    $computers = Get-Content -LiteralPath $ListPath
}
else {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw 'Модуль ActiveDirectory не найден. Установите RSAT или укажите -ComputerName / -ListPath.'
    }
    Import-Module ActiveDirectory

    $adParams = @{
        Filter     = { OperatingSystem -like $OperatingSystem }
        Properties = 'OperatingSystem', 'Enabled'
    }
    if ($SearchBase) { $adParams.SearchBase = $SearchBase }
    if ($Credential) { $adParams.Credential = $Credential }

    $adComputers = Get-ADComputer @adParams
    if (-not $IncludeDisabled) { $adComputers = $adComputers | Where-Object { $_.Enabled } }
    $computers = $adComputers | Select-Object -ExpandProperty Name
}

$computers = @($computers | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
if ($computers.Count -eq 0) { throw 'Список компьютеров пуст.' }
Write-Host "Компьютеров к опросу: $($computers.Count)" -ForegroundColor Cyan
#endregion

#region Рабочий блок (выполняется в отдельном runspace для каждого компьютера)
$worker = {
    param([string]$Computer, $Cred, [int]$Timeout, [bool]$SkipUsb)

    function Convert-DiskSerial([string]$Serial) {
        if (-not $Serial) { return '' }
        $s = $Serial.Trim()
        # Некоторые драйверы (SATA/IDE) отдают серийник как 40 HEX-символов с переставленными байтами
        if ($s -match '^[0-9A-Fa-f]{40}$') {
            $text = -join (0..19 | ForEach-Object { [char][Convert]::ToInt32($s.Substring($_ * 2, 2), 16) })
            $swapped = -join (0..9 | ForEach-Object { "$($text[$_ * 2 + 1])$($text[$_ * 2])" })
            if ($swapped -match '^[\x20-\x7E]+$') { return $swapped.Trim() }
        }
        return $s.TrimEnd('.').Trim()
    }

    $result = [pscustomobject]@{ Computer = $Computer; Status = 'Error'; Message = ''; Disks = @() }
    $session = $null
    try {
        if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
            $result.Status = 'Offline'
            $result.Message = 'Нет ответа на ping'
            return $result
        }

        # Для локального компьютера учетные данные не передаются
        $isLocal = ($Computer -eq $env:COMPUTERNAME) -or ($Computer -like "$env:COMPUTERNAME.*")
        $sessParams = @{ ComputerName = $Computer; OperationTimeoutSec = $Timeout; ErrorAction = 'Stop' }
        if ($Cred -and -not $isLocal) { $sessParams.Credential = $Cred }

        try {
            $session = New-CimSession @sessParams
            $null = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem -ErrorAction Stop
        }
        catch {
            if ($session) { Remove-CimSession $session -ErrorAction SilentlyContinue }
            $sessParams.SessionOption = New-CimSessionOption -Protocol Dcom
            $session = New-CimSession @sessParams
        }

        $cs = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem
        $os = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem
        $driveList = @(Get-CimInstance -CimSession $session -ClassName Win32_DiskDrive)

        $media = @{}
        try {
            Get-CimInstance -CimSession $session -ClassName Win32_PhysicalMedia -ErrorAction Stop |
                ForEach-Object { if ($_.Tag) { $media[$_.Tag.ToUpper()] = $_.SerialNumber } }
        } catch { }

        $physical = @{}
        try {
            Get-CimInstance -CimSession $session -Namespace 'root\Microsoft\Windows\Storage' `
                -ClassName MSFT_PhysicalDisk -ErrorAction Stop |
                ForEach-Object { $physical["$($_.DeviceId)"] = $_ }
        } catch { }   # на Windows 7 / Server 2008 R2 класса нет

        $disks = foreach ($d in $driveList) {
            if ($SkipUsb -and $d.InterfaceType -eq 'USB') { continue }

            $idx = "$($d.Index)"
            $pd = $physical[$idx]
            $mediaSerial = $media["\\.\PHYSICALDRIVE$idx"]

            $serial = ''
            foreach ($candidate in @(
                    $(if ($pd) { $pd.SerialNumber }), $d.SerialNumber, $mediaSerial)) {
                $c = Convert-DiskSerial $candidate
                if ($c) { $serial = $c; break }
            }

            $mediaType = ''
            if ($pd) {
                $mediaType = switch ([int]$pd.MediaType) { 3 { 'HDD' } 4 { 'SSD' } 5 { 'SCM' } default { '' } }
            }

            [pscustomobject][ordered]@{
                ComputerName    = $Computer
                OperatingSystem = $os.Caption
                ComputerModel   = ("$($cs.Manufacturer) $($cs.Model)").Trim()
                DiskNumber      = $d.Index
                Model           = ("$($d.Model)").Trim()
                SerialNumber    = $serial
                UniqueId        = if ($pd) { "$($pd.UniqueId)".Trim() } else { '' }
                PNPDeviceID     = $d.PNPDeviceID
                MediaType       = $mediaType
                InterfaceType   = $d.InterfaceType
                SizeGB          = [math]::Round([double]$d.Size / 1GB, 1)
                Firmware        = $d.FirmwareRevision
                Health          = if ($pd) { "$($pd.HealthStatus)" } else { "$($d.Status)" }
                Partitions      = $d.Partitions
                CollectedAt     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
        }

        $result.Disks = @($disks | Where-Object { $_ })
        $result.Status = 'OK'
        $result.Message = "Дисков: $($result.Disks.Count)"
    }
    catch {
        $result.Status = 'Error'
        $result.Message = $_.Exception.Message
    }
    finally {
        if ($session) { Remove-CimSession $session -ErrorAction SilentlyContinue }
    }
    return $result
}
#endregion

#region Параллельный опрос
$pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
$pool.Open()

$jobs = foreach ($c in $computers) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    $null = $ps.AddScript($worker).AddArgument($c).AddArgument($Credential).AddArgument($TimeoutSec).AddArgument([bool]$ExcludeUsb)
    [pscustomobject]@{ Computer = $c; PS = $ps; Handle = $ps.BeginInvoke() }
}

$results = New-Object System.Collections.Generic.List[object]
$pending = New-Object System.Collections.Generic.List[object]
$pending.AddRange(@($jobs))
$total = $pending.Count

while ($pending.Count -gt 0) {
    foreach ($j in @($pending | Where-Object { $_.Handle.IsCompleted })) {
        try {
            $out = $j.PS.EndInvoke($j.Handle)
            if ($out -and $out.Count -gt 0) { $results.Add($out[$out.Count - 1]) }
            else { $results.Add([pscustomobject]@{ Computer = $j.Computer; Status = 'Error'; Message = 'Пустой ответ'; Disks = @() }) }
        }
        catch {
            $results.Add([pscustomobject]@{ Computer = $j.Computer; Status = 'Error'; Message = $_.Exception.Message; Disks = @() })
        }
        finally {
            $j.PS.Dispose()
            $null = $pending.Remove($j)
        }
    }
    $done = $total - $pending.Count
    Write-Progress -Activity 'Опрос компьютеров домена' -Status "$done из $total" -PercentComplete ($done / $total * 100)
    if ($pending.Count -gt 0) { Start-Sleep -Milliseconds 300 }
}
Write-Progress -Activity 'Опрос компьютеров домена' -Completed
$pool.Close()
$pool.Dispose()
#endregion

#region Отчеты
if (-not (Test-Path -LiteralPath $OutputDir)) { $null = New-Item -ItemType Directory -Path $OutputDir -Force }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath = Join-Path $OutputDir "DiskInventory_$stamp.csv"
$htmlPath = Join-Path $OutputDir "DiskInventory_$stamp.html"
$errPath = Join-Path $OutputDir "DiskInventory_${stamp}_errors.csv"
$encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

$allDisks = @($results | ForEach-Object { $_.Disks } | Where-Object { $_ } | Sort-Object ComputerName, DiskNumber)
$failed = @($results | Where-Object { $_.Status -ne 'OK' } | Sort-Object Computer |
        Select-Object @{n = 'ComputerName'; e = { $_.Computer } }, Status, Message)
$okCount = @($results | Where-Object { $_.Status -eq 'OK' }).Count

if ($allDisks.Count -gt 0) {
    $allDisks | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ';' -Encoding $encoding
}
if ($failed.Count -gt 0) {
    $failed | Export-Csv -Path $errPath -NoTypeInformation -Delimiter ';' -Encoding $encoding
}

$css = @'
<style>
body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:20px;color:#222}
h1{font-size:20px} h2{font-size:16px;margin-top:28px}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;white-space:nowrap}
th{background:#2f5d8a;color:#fff;position:sticky;top:0}
tr:nth-child(even){background:#f3f6fa}
.summary td{border:none;padding:2px 12px 2px 0}
</style>
'@
$summaryHtml = @"
<h1>Инвентаризация жестких дисков домена</h1>
<table class="summary">
<tr><td>Дата отчета:</td><td>$(Get-Date -Format 'dd.MM.yyyy HH:mm')</td></tr>
<tr><td>Компьютеров опрошено:</td><td>$($results.Count)</td></tr>
<tr><td>Успешно:</td><td>$okCount</td></tr>
<tr><td>С ошибкой / недоступны:</td><td>$($failed.Count)</td></tr>
<tr><td>Всего дисков:</td><td>$($allDisks.Count)</td></tr>
</table>
"@

$disksHtml = if ($allDisks.Count -gt 0) {
    $allDisks | Select-Object ComputerName, DiskNumber, Model, SerialNumber, UniqueId, PNPDeviceID,
    MediaType, InterfaceType, SizeGB, Firmware, Health, OperatingSystem |
        ConvertTo-Html -Fragment | Out-String
} else { '<p>Данные не получены.</p>' }

$failedHtml = if ($failed.Count -gt 0) {
    '<h2>Недоступные компьютеры</h2>' + ($failed | ConvertTo-Html -Fragment | Out-String)
} else { '' }

ConvertTo-Html -Title 'Инвентаризация дисков' -Head "<meta charset=`"utf-8`">$css" `
    -Body "$summaryHtml<h2>Диски</h2>$disksHtml$failedHtml" |
    Out-File -FilePath $htmlPath -Encoding utf8
#endregion

Write-Host ''
Write-Host "Успешно опрошено: $okCount из $($results.Count); дисков найдено: $($allDisks.Count)" -ForegroundColor Green
if ($allDisks.Count -gt 0) { Write-Host "CSV : $csvPath" }
Write-Host "HTML: $htmlPath"
if ($failed.Count -gt 0) {
    Write-Host "Ошибки ($($failed.Count)): $errPath" -ForegroundColor Yellow
}
