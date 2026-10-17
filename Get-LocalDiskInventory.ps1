<#
.SYNOPSIS
    Собирает данные о жестких дисках локального компьютера (модель, серийный и
    электронный номера) и формирует отчеты CSV и HTML.

.DESCRIPTION
    Локальная версия Get-DomainDiskInventory.ps1: те же источники и те же поля отчета,
    но без Active Directory и удаленных подключений.
      - Win32_DiskDrive              : модель, серийный номер, PNPDeviceID, интерфейс, размер
      - Win32_PhysicalMedia          : серийный номер носителя (запасной источник)
      - MSFT_PhysicalDisk (Storage)  : серийный номер, UniqueId (электронный идентификатор),
                                       тип носителя SSD/HDD, состояние, прошивка

    Результат выводится в консоль таблицей и сохраняется в файлы:
      DiskInventory_<компьютер>_<дата>.csv   - разделитель ";", UTF-8 с BOM (открывается в Excel)
      DiskInventory_<компьютер>_<дата>.html

    Серийный номер: берется из MSFT_PhysicalDisk, иначе из Win32_DiskDrive, иначе из
    Win32_PhysicalMedia. Если драйвер отдает его в виде HEX-строки, она декодируется.
    Электронный номер: UniqueId (WWN/NAA/EUI/SCSI-идентификатор диска) и PNPDeviceID.

.PARAMETER ExcludeUsb
    Не включать диски с интерфейсом USB.

.PARAMETER OutputDir
    Папка для отчетов. По умолчанию - подпапка Reports рядом со скриптом
    (если скрипт запущен не из файла - в текущей папке).

.PARAMETER NoReport
    Только вывести таблицу в консоль, файлы не создавать.

.EXAMPLE
    .\Get-LocalDiskInventory.ps1

.EXAMPLE
    .\Get-LocalDiskInventory.ps1 -ExcludeUsb -OutputDir D:\Reports

.EXAMPLE
    .\Get-LocalDiskInventory.ps1 -NoReport

.NOTES
    Требуется: Windows, PowerShell 5.1+. Запускайте от имени администратора:
    без повышенных прав класс MSFT_PhysicalDisk (UniqueId, тип SSD/HDD) может быть недоступен.
#>
[CmdletBinding()]
param(
    [switch]$ExcludeUsb,
    [string]$OutputDir,
    [switch]$NoReport
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# $PSScriptRoot пуст, если код запущен не из файла (вставка в консоль, F8 в ISE, dot-sourcing)
if (-not $OutputDir) {
    $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputDir = Join-Path $baseDir 'Reports'
}

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

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning 'Скрипт запущен без прав администратора: UniqueId и тип носителя (SSD/HDD) могут быть недоступны.'
}

#region Сбор данных
$computer = $env:COMPUTERNAME
$cs = Get-CimInstance -ClassName Win32_ComputerSystem
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$driveList = @(Get-CimInstance -ClassName Win32_DiskDrive)

$media = @{}
try {
    Get-CimInstance -ClassName Win32_PhysicalMedia -ErrorAction Stop |
        ForEach-Object { if ($_.Tag) { $media[$_.Tag.ToUpper()] = $_.SerialNumber } }
} catch { }

$physical = @{}
try {
    Get-CimInstance -Namespace 'root\Microsoft\Windows\Storage' -ClassName MSFT_PhysicalDisk -ErrorAction Stop |
        ForEach-Object { $physical["$($_.DeviceId)"] = $_ }
} catch { }   # на Windows 7 / Server 2008 R2 класса нет

$disks = foreach ($d in $driveList) {
    if ($ExcludeUsb -and $d.InterfaceType -eq 'USB') { continue }

    $idx = "$($d.Index)"
    $pd = $physical[$idx]
    $mediaSerial = $media["\\.\PHYSICALDRIVE$idx"]

    $serial = ''
    foreach ($candidate in @($(if ($pd) { $pd.SerialNumber }), $d.SerialNumber, $mediaSerial)) {
        $c = Convert-DiskSerial $candidate
        if ($c) { $serial = $c; break }
    }

    $mediaType = ''
    if ($pd) {
        $mediaType = switch ([int]$pd.MediaType) { 3 { 'HDD' } 4 { 'SSD' } 5 { 'SCM' } default { '' } }
    }

    [pscustomobject][ordered]@{
        ComputerName    = $computer
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
$disks = @($disks | Where-Object { $_ } | Sort-Object DiskNumber)
#endregion

#region Вывод
Write-Host "Компьютер: $computer ($($os.Caption)); дисков найдено: $($disks.Count)" -ForegroundColor Cyan
if ($disks.Count -eq 0) {
    Write-Warning 'Диски не найдены.'
    return
}
$disks | Format-Table DiskNumber, Model, SerialNumber, UniqueId, MediaType, InterfaceType, SizeGB, Health -AutoSize -Wrap
#endregion

#region Отчеты
if (-not $NoReport) {
    if (-not (Test-Path -LiteralPath $OutputDir)) { $null = New-Item -ItemType Directory -Path $OutputDir -Force }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath = Join-Path $OutputDir "DiskInventory_${computer}_$stamp.csv"
    $htmlPath = Join-Path $OutputDir "DiskInventory_${computer}_$stamp.html"
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

    $disks | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ';' -Encoding $encoding

    $css = @'
<style>
body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:20px;color:#222}
h1{font-size:20px} h2{font-size:16px;margin-top:28px}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;white-space:nowrap}
th{background:#2f5d8a;color:#fff}
tr:nth-child(even){background:#f3f6fa}
.summary td{border:none;padding:2px 12px 2px 0}
</style>
'@
    $summaryHtml = @"
<h1>Жесткие диски компьютера $([System.Net.WebUtility]::HtmlEncode($computer))</h1>
<table class="summary">
<tr><td>Дата отчета:</td><td>$(Get-Date -Format 'dd.MM.yyyy HH:mm')</td></tr>
<tr><td>Операционная система:</td><td>$([System.Net.WebUtility]::HtmlEncode($os.Caption))</td></tr>
<tr><td>Модель компьютера:</td><td>$([System.Net.WebUtility]::HtmlEncode(("$($cs.Manufacturer) $($cs.Model)").Trim()))</td></tr>
<tr><td>Всего дисков:</td><td>$($disks.Count)</td></tr>
</table>
"@
    $disksHtml = $disks | Select-Object DiskNumber, Model, SerialNumber, UniqueId, PNPDeviceID,
    MediaType, InterfaceType, SizeGB, Firmware, Health, Partitions |
        ConvertTo-Html -Fragment | Out-String

    ConvertTo-Html -Title "Диски $computer" -Head "<meta charset=`"utf-8`">$css" `
        -Body "$summaryHtml<h2>Диски</h2>$disksHtml" |
        Out-File -FilePath $htmlPath -Encoding utf8

    Write-Host "CSV : $csvPath"
    Write-Host "HTML: $htmlPath"
}
#endregion
