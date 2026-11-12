<#
.SYNOPSIS
    Поиск прав доступа конкретного пользователя по отчёту FileAccessAudit.ps1.

.DESCRIPTION
    Читает CSV-отчёт AccessMatrix_*.csv, созданный скриптом FileAccessAudit.ps1,
    и показывает все папки, к которым у пользователя есть доступ — напрямую
    или через группы. Результат выводится в консоль и сохраняется в CSV.

    ВНИМАНИЕ: этот файл генерируется автоматически при каждом запуске
    FileAccessAudit.ps1. Если вы изменили его вручную, изменения будут
    потеряны при следующем запуске аудита.

.PARAMETER UserName
    Логин (sAMAccountName) или часть ФИО. Поиск идёт по подстроке.

.PARAMETER ReportPath
    Путь к конкретному CSV-отчёту. Если не указан, берётся самый свежий
    AccessMatrix_*.csv из каталога .\log рядом со скриптом.

.EXAMPLE
    .\SearchUserAccess.ps1 -UserName ivanov

.EXAMPLE
    .\SearchUserAccess.ps1 -UserName "Иванов" -ReportPath .\log\AccessMatrix_20260101_120000.csv
#>
[CmdletBinding()]
param(
    [string]$UserName,
    [string]$ReportPath
)

# Каталог с отчётами аудита. По умолчанию — подкаталог log рядом со скриптом.
# Если $PSScriptRoot не определён (запуск через Invoke-Expression), берём текущий каталог.
if ($PSScriptRoot) {
    $ReportDir = Join-Path $PSScriptRoot 'log'
}
else {
    $ReportDir = Join-Path (Get-Location).Path 'log'
}

if (-not $UserName) {
    $UserName = Read-Host 'Введите имя пользователя для поиска'
}

# Если конкретный CSV не указан — берём самый свежий отчёт AccessMatrix_*.csv.
if (-not $ReportPath) {
    $latestReport = Get-ChildItem -Path $ReportDir -Filter 'AccessMatrix_*.csv' -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1

    if (-not $latestReport) {
        Write-Host 'Отчеты не найдены. Сначала запустите FileAccessAudit.ps1.' -ForegroundColor Red
        exit 1
    }

    $ReportPath = $latestReport.FullName
}

$accessData = Import-Csv -Path $ReportPath -Delimiter ';'

# Поиск по логину или по ФИО (частичное совпадение, без учёта регистра).
$userAccess = $accessData | Where-Object {
    $_.Пользователь -like "*$UserName*" -or $_.ПолноеИмя -like "*$UserName*"
}

if (-not $userAccess) {
    Write-Host "Права доступа для '$UserName' не найдены" -ForegroundColor Yellow
    exit 0
}

# Считаем, сколько доступов выдано напрямую, а сколько — через группы.
# @() нужен, чтобы .Count работал и при единственной найденной строке.
$accessCount = @($userAccess).Count
$directCount = @($userAccess | Where-Object { $_.ИсточникДоступа -eq 'Прямой доступ' }).Count
$groupCount = @($userAccess | Where-Object { $_.ИсточникДоступа -eq 'Через группу' }).Count

Write-Host "Найдено прав доступа для '$UserName': $accessCount" -ForegroundColor Green
Write-Host "Прямых доступов: $directCount"
Write-Host "Доступов через группы: $groupCount"
Write-Host ""

$userAccess | Format-Table Путь, ИсточникДоступа, Группа, Права, ТипДоступа, Наследование -AutoSize

# Сохраняем результат поиска рядом с исходным отчётом.
# Недопустимые в имени файла символы заменяем на "_".
$safeUserName = $UserName -replace '[\\/:*?"<>|]', '_'
$resultDir = Split-Path -Parent $ReportPath
$userReportPath = Join-Path $resultDir ("UserAccess_{0}_{1}.csv" -f $safeUserName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
$userAccess | Export-Csv -Path $userReportPath -Encoding UTF8 -NoTypeInformation -Delimiter ';'
Write-Host "Отчет сохранен: $userReportPath" -ForegroundColor Green
