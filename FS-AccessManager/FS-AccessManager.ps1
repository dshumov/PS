#Requires -Version 5.1
<#
.SYNOPSIS
    FS-AccessManager - управление доступом к папкам файлового сервера.

.DESCRIPTION
    Утилита с графическим интерфейсом (WinForms) для администраторов Active Directory.

    Возможности:
      * выбор существующей папки на файловом сервере или создание новой;
      * проверка наличия групп AD FS_<папка>_RO / FS_<папка>_RW / FS_<папка>_FA;
      * при необходимости создание OU с именем папки (внутри выбранной OU), в которой создаются группы;
      * создание отсутствующих групп в указанном подразделении (OU):
            Group category : Security
            Group scope    : Domain Local
        с описаниями:
            FS_<папка>_RO - Группа доступа к <папка> (Только чтение)
            FS_<папка>_RW - Группа доступа к <папка> (Чтение и Запись)
            FS_<папка>_FA - Группа доступа к <папка> (Полный доступ)
      * назначение этих групп на папку с наследованием на подпапки и файлы:
            RO -> ReadAndExecute (чтение и выполнение)
            RW -> Modify (изменение)
            FA -> FullControl (полный доступ)
      * добавление пользователей в группы, просмотр и удаление участников.

.PARAMETER DomainController
    Контроллер домена (имя или IP-адрес). Если не задан - выбирается в окне.

.PARAMETER FileServer
    Имя или IP-адрес файлового сервера.

.PARAMETER RootPath
    Корневая папка на файловом сервере (UNC-путь или локальный путь).

.PARAMETER TargetOU
    DN подразделения, в котором создаются группы.
    Пример: "OU=FileShares,OU=Resources,DC=example,DC=local"

.EXAMPLE
    .\FS-AccessManager.ps1

.EXAMPLE
    .\FS-AccessManager.ps1 -DomainController dc01.example.local -FileServer fileserver -RootPath \\fileserver\Shared -TargetOU "OU=FileShares,DC=example,DC=local"

.NOTES
    Требования:
      * Windows PowerShell 5.1 или PowerShell 7.x;
      * модуль ActiveDirectory (RSAT): Install-WindowsFeature RSAT-AD-PowerShell
        либо Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0;
      * права на создание групп в OU и на изменение NTFS-разрешений на файловом сервере;
      * доступ к папке на файловом сервере под текущей учётной записью
        (файловые операции всегда выполняются от текущей учётной записи).

    Версия: 1.2.
    Файл сохранён в кодировке UTF-8 с BOM. Не пересохраняйте его в ANSI/Windows-1251.
#>

[CmdletBinding()]
param(
    [string]$DomainController,
    [string]$FileServer,
    [string]$RootPath,
    [string]$TargetOU
)

# ============================================================================
#  КАК УСТРОЕН СКРИПТ (карта разделов)
#   1. Общие данные и состояние ....... настройки, пути, флаги, служебные данные
#   2. Конструкторы интерфейса ........ фабрики WinForms-контролов
#   3. Журнал и сообщения ............. вывод в журнал, диалоги, строка состояния
#   4. Работа с путями ................ нормализация, проверки, имена папок
#   5. Настройки ...................... JSON рядом со скриптом (или в LOCALAPPDATA)
#   6. Active Directory ............... группы, OU, DN, параметры подключения
#   7. Основные операции .............. DC, папки, статусы групп, участники
#   8. Вспомогательные функции ........ UNC-пути, разбор DN, дерево OU, креды, шары
#   9. Диалоги ........................ папка, OU, сервер, пользователи
#  10. Построение главного окна ....... компоновка и обработчики событий
#  11. Запуск ......................... импорт AD, окно, цикл сообщений
#
#  Поток работы: выбор папки -> проверка групп AD -> (создание OU и групп) ->
#  назначение NTFS-прав -> добавление пользователей в группы.
# ============================================================================

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Эти вызовы допустимы только до создания первого окна в процессе.
# При запуске скрипта из уже открытой консоли PowerShell окно уже может существовать,
# поэтому оборачиваем в try/catch и не прерываем запуск.
try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$script:NL = [Environment]::NewLine


#region ---------- Общие данные ----------
# Здесь хранятся: путь к файлу настроек, объект состояния $script:State (домен,
# контроллер, сервер, корневой путь, OU, учётные данные, флаги) и служебные данные.

$script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:SettingsFile = Join-Path $script:ScriptDir 'FS-AccessManager.settings.json'
$script:AppTitle = 'FS-AccessManager - доступ к папкам файлового сервера'

$script:State = [pscustomobject]@{
    Domain             = ''
    DomainDN           = ''
    NetBIOS            = ''
    DC                 = ''
    FileServer         = ''
    RootPath           = ''
    TargetOU           = ''
    CurrentPath        = ''
    SelectedPath       = ''
    Credential         = $null
    ADReady            = $false
    DisableInheritance = $false
    CreateFolderOU     = $true
}

$script:DCList = @()
$script:DCMap = @{}
$script:ServerHistory = @()
$script:SuppressEvents = $false
$script:ModuleError = ''
$script:CredentialPromptDeclined = $false
$script:LogBox = $null
$script:StatusLabel = $null
$script:Form = $null

#endregion


#region ---------- Элементы интерфейса (конструкторы) ----------
# Небольшие фабрики WinForms-контролов: чтобы не дублировать код создания
# меток, кнопок, полей, списков и флажков.

# Конструктор Label: текст, координаты, размер, выравнивание.
function New-UILabel {
    param(
        $Parent,
        [string]$Text = '',
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 120,
        [int]$Height = 20,
        [switch]$Bold
    )
    $c = New-Object System.Windows.Forms.Label
    $c.Text = $Text
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($Width, $Height)
    $c.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $c.AutoEllipsis = $true
    if ($Bold) { $c.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }
    $Parent.Controls.Add($c)
    return $c
}

# Конструктор кнопки с автоматической подпиской на событие Click.
function New-UIButton {
    param(
        $Parent,
        [string]$Text,
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 110,
        [int]$Height = 27,
        [scriptblock]$OnClick
    )
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($Width, $Height)
    $b.UseVisualStyleBackColor = $true
    if ($OnClick) { $b.Add_Click($OnClick) }
    $Parent.Controls.Add($b)
    return $b
}

# Конструктор TextBox (включая режим "только чтение").
function New-UITextBox {
    param(
        $Parent,
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 300,
        [int]$Height = 25,
        [switch]$ReadOnly
    )
    $c = New-Object System.Windows.Forms.TextBox
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($Width, $Height)
    $c.ReadOnly = [bool]$ReadOnly
    if ($ReadOnly) { $c.BackColor = [System.Drawing.Color]::WhiteSmoke }
    $Parent.Controls.Add($c)
    return $c
}

# Конструктор ComboBox: редактируемый (с автодополнением) или выбор из списка.
function New-UIComboBox {
    param(
        $Parent,
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 300,
        [switch]$Editable
    )
    $c = New-Object System.Windows.Forms.ComboBox
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($Width, 25)
    if ($Editable) {
        $c.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
        $c.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
        $c.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
    } else {
        $c.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    }
    $Parent.Controls.Add($c)
    return $c
}

# Конструктор CheckBox.
function New-UICheckBox {
    param(
        $Parent,
        [string]$Text,
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 250,
        [int]$Height = 24
    )
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($Width, $Height)
    $c.UseVisualStyleBackColor = $true
    $Parent.Controls.Add($c)
    return $c
}

#endregion


#region ---------- Журнал и сообщения ----------
# Всё, что связано с общением с пользователем: цветной журнал, строка состояния,
# окна ошибок/предупреждений и подтверждений, склейка многострочных сообщений.

# Добавляет строку в журнал: время, уровень (Info/Ok/Warn/Error) и цвет.
function Write-UiLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error')]
        [string]$Level = 'Info'
    )
    if (-not $script:LogBox) { return }
    $stamp = (Get-Date).ToString('HH:mm:ss')
    $prefix = '    '
    switch ($Level) {
        'Ok'    { $prefix = 'OK  ' }
        'Warn'  { $prefix = 'WARN' }
        'Error' { $prefix = 'ERR ' }
    }
    $color = [System.Drawing.Color]::Black
    switch ($Level) {
        'Ok'    { $color = [System.Drawing.Color]::FromArgb(0, 110, 0) }
        'Warn'  { $color = [System.Drawing.Color]::FromArgb(170, 90, 0) }
        'Error' { $color = [System.Drawing.Color]::FromArgb(180, 0, 0) }
    }
    try {
        $script:LogBox.SelectionStart = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0
        $script:LogBox.SelectionColor = $color
        $script:LogBox.AppendText($stamp + '  ' + $prefix + '  ' + $Message + [Environment]::NewLine)
        $script:LogBox.SelectionColor = $script:LogBox.ForeColor
        $script:LogBox.ScrollToCaret()
        $script:LogBox.Update()
    } catch { }
}

# Обновляет подсказку под галочкой "Создавать OU с именем папки".
function Update-CreateOuHint {
    if (-not $script:lblCreateOuHint) { return }
    if ($script:State.CreateFolderOU) {
        $script:lblCreateOuHint.Text = 'группы создаются в OU=ИмяПапки внутри выбранной OU'
    } else {
        $script:lblCreateOuHint.Text = 'группы создаются непосредственно в выбранной OU'
    }
}

# Показывает короткий текст в строке состояния внизу окна.
function Set-Status {
    param([string]$Text)
    if ($script:StatusLabel) {
        $script:StatusLabel.Text = $Text
        $script:StatusLabel.Update()
    }
}

# Склеивает массив строк через перевод строки (для многострочных сообщений).
function Join-Lines {
    param([string[]]$Lines)
    if (-not $Lines) { return '' }
    return ($Lines -join [Environment]::NewLine)
}

# Модальное окно с ошибкой.
function Show-ErrorBox {
    param([string]$Message, [string]$Title = 'Ошибка')
    [void][System.Windows.Forms.MessageBox]::Show($script:Form, $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Модальное окно с предупреждением.
function Show-WarnBox {
    param([string]$Message, [string]$Title = 'Внимание')
    [void][System.Windows.Forms.MessageBox]::Show($script:Form, $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
}

# Диалог подтверждения Да/Нет; возвращает $true при выборе "Да".
function Confirm-Action {
    param([string]$Message, [string]$Title = 'Подтверждение')
    $r = [System.Windows.Forms.MessageBox]::Show($script:Form, $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

#endregion


#region ---------- Работа с путями ----------
# Нормализация путей, проверка "путь внутри корня", валидация имён папок
# и проверка совместимости имени папки с именем группы AD.

# Приводит путь к единому виду: слэши, кавычки, хвостовой "", UNC-корни.
function Normalize-Path {
    param([string]$Path)
    if (-not $Path) { return '' }
    $p = $Path.Trim()
    if ($p.StartsWith('"') -and $p.EndsWith('"') -and $p.Length -ge 2) {
        $p = $p.Substring(1, $p.Length - 2)
    }
    $p = $p -replace '/', '\'
    $p = $p.TrimEnd(' ')
    if ($p -match '^[A-Za-z]:$') { $p = $p + '\' }
    if ($p -match '^\\\\[^\\]+$') { return $p }
    if ($p -match '^[A-Za-z]:\\$') { return $p }
    while ($p.Length -gt 3 -and $p.EndsWith('\')) {
        $p = $p.Substring(0, $p.Length - 1)
    }
    return $p
}

# Проверяет, что путь находится внутри корневой папки (защита от выхода за корень).
function Test-PathUnderRoot {
    param([string]$Path, [string]$Root)
    $p = Normalize-Path $Path
    $r = Normalize-Path $Root
    if (-not $p -or -not $r) { return $false }
    if ($p.Equals($r, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    $prefix = $r.TrimEnd('\') + '\'
    return $p.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

# Валидация имени новой папки: запрещённые символы, длина, занятость имени.
function Test-NewFolderName {
    param([string]$Name, [string]$ParentPath)
    $n = $Name.Trim()
    if (-not $n) { return 'Введите имя папки.' }
    if ($n -eq '.' -or $n -eq '..') { return 'Недопустимое имя папки.' }
    if ($n -match '[\\/:*?"<>|]') { return 'Имя содержит недопустимые символы: \ / : * ? " < > |' }
    if ($n -match '[\[\];=,+]') { return 'Имя содержит символы, недопустимые в sAMAccountName группы AD: [ ] ; = , +' }
    if ($n -match '[\. ]$') { return 'Имя не должно заканчиваться точкой или пробелом.' }
    if ($n.Length -gt 56) { return 'Имя папки длиннее 56 символов - имя группы AD превысит 64 символа.' }
    if ($ParentPath) {
        $full = Join-Path $ParentPath $n
        if (Test-Path -LiteralPath $full) { return 'Папка с таким именем уже существует.' }
    }
    return $null
}

# Проверяет, что из имени папки получится корректное имя группы AD (CN/sAMAccountName).
function Assert-GroupNameCompatibility {
    param([string]$FolderName)
    if ($FolderName -match '[\[\];=,+]') {
        throw ('Имя папки "' + $FolderName + '" содержит символы, недопустимые в sAMAccountName группы AD: [ ] ; = , +.' + [Environment]::NewLine + 'Переименуйте папку или создайте группы вручную.')
    }
    $sample = 'FS_' + $FolderName + '_RO'
    if ($sample.Length -gt 64) {
        throw ('Имя группы "' + $sample + '" длиннее 64 символов - AD не позволит создать такой объект. Сократите имя папки.')
    }
    if ($sample.Length -gt 20) {
        Write-UiLog ('Внимание: имя группы "' + $sample + '" длиннее 20 символов (ограничение sAMAccountName). Некоторые инструменты могут отображать логин группы усечённым.') 'Warn'
    }
}

#endregion


#region ---------- Настройки ----------
# Чтение и запись JSON-настроек (рядом со скриптом), история файловых серверов.
# Пароли и учётные данные в настройках не сохраняются.

# Читает JSON с настройками (рядом со скриптом или в %LOCALAPPDATA%).
function Read-SettingsFile {
    $candidates = @($script:SettingsFile, (Join-Path $env:LOCALAPPDATA 'FS-AccessManager\settings.json'))
    foreach ($path in $candidates) {
        if (-not $path) { continue }
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
            if (-not $raw) { continue }
            return ($raw | ConvertFrom-Json)
        } catch {
            Write-UiLog ('Не удалось прочитать настройки "' + $path + '": ' + $_.Exception.Message) 'Warn'
        }
    }
    return $null
}

# Применяет настройки и параметры запуска к состоянию и элементам окна.
function Initialize-Settings {
    $cfg = $null
    try { $cfg = Read-SettingsFile } catch { $cfg = $null }

    if ($cfg) {
        if ($cfg.FileServers) {
            $script:ServerHistory = @($cfg.FileServers | ForEach-Object { [string]$_ })
        }
        if ($cfg.DomainController) { $script:State.DC = [string]$cfg.DomainController }
        if ($cfg.FileServer) { $script:State.FileServer = [string]$cfg.FileServer }
        if ($cfg.RootPath) { $script:State.RootPath = Normalize-Path ([string]$cfg.RootPath) }
        if ($cfg.TargetOU) { $script:State.TargetOU = [string]$cfg.TargetOU }
        if ($cfg.DisableInheritance) { $script:State.DisableInheritance = [bool]$cfg.DisableInheritance }
        if ($null -ne $cfg.CreateFolderOU) { $script:State.CreateFolderOU = [bool]$cfg.CreateFolderOU }
    }

    if ($DomainController) { $script:State.DC = $DomainController }
    if ($FileServer) { $script:State.FileServer = $FileServer }
    if ($RootPath) { $script:State.RootPath = Normalize-Path $RootPath }
    if ($TargetOU) { $script:State.TargetOU = $TargetOU }

    $script:SuppressEvents = $true
    try {
        if ($script:State.FileServer) { $script:cmbFS.Text = $script:State.FileServer }
        foreach ($s in $script:ServerHistory) {
            if ($s -and -not $script:cmbFS.Items.Contains($s)) { [void]$script:cmbFS.Items.Add($s) }
        }
        if ($script:State.RootPath) { $script:txtRoot.Text = $script:State.RootPath }
        if ($script:State.TargetOU) { $script:txtOU.Text = $script:State.TargetOU }
        $script:chkInherit.Checked = [bool]$script:State.DisableInheritance
        $script:chkCreateFolderOU.Checked = [bool]$script:State.CreateFolderOU
        Update-CreateOuHint
        if ($script:State.DC) { $script:cmbDC.Text = $script:State.DC }
    } finally {
        $script:SuppressEvents = $false
    }
}

# Сохраняет текущие параметры в JSON (пароли не сохраняются).
function Save-Settings {
    try {
        $fsText = $script:State.FileServer
        if ($script:cmbFS) { $fsText = $script:cmbFS.Text.Trim() }
        $rootText = $script:State.RootPath
        if ($script:txtRoot) { $rootText = $script:txtRoot.Text.Trim() }
        $ouText = $script:State.TargetOU
        if ($script:txtOU) { $ouText = $script:txtOU.Text.Trim() }
        $inherit = $script:State.DisableInheritance
        if ($script:chkInherit) { $inherit = [bool]$script:chkInherit.Checked }
        $createFolderOu = $script:State.CreateFolderOU
        if ($script:chkCreateFolderOU) { $createFolderOu = [bool]$script:chkCreateFolderOU.Checked }

        $data = [ordered]@{
            Version            = 1
            DomainController   = $script:State.DC
            FileServer         = $fsText
            RootPath           = $rootText
            TargetOU           = $ouText
            DisableInheritance = $inherit
            CreateFolderOU     = $createFolderOu
            FileServers        = @($script:ServerHistory)
            SavedAt            = (Get-Date).ToString('s')
        }
        $json = $data | ConvertTo-Json -Depth 4
        $paths = @($script:SettingsFile, (Join-Path $env:LOCALAPPDATA 'FS-AccessManager\settings.json'))
        foreach ($p in $paths) {
            try {
                $dir = Split-Path -Parent $p
                if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                    [void][System.IO.Directory]::CreateDirectory($dir)
                }
                [System.IO.File]::WriteAllText($p, $json, (New-Object System.Text.UTF8Encoding($false)))
                return
            } catch { }
        }
    } catch { }
}

# Добавляет файловый сервер в историю выпадающего списка.
function Add-ServerHistory {
    param([string]$Server)
    if (-not $Server) { return }
    $list = @($script:ServerHistory)
    if ($list -contains $Server) { return }
    $list = @($Server) + @($list)
    $script:ServerHistory = @($list | Select-Object -First 10)
    if ($script:cmbFS) {
        $keepText = $script:cmbFS.Text
        $script:SuppressEvents = $true
        try {
            $script:cmbFS.BeginUpdate()
            try {
                $script:cmbFS.Items.Clear()
                foreach ($s in $script:ServerHistory) { [void]$script:cmbFS.Items.Add($s) }
            } finally {
                $script:cmbFS.EndUpdate()
            }
            $script:cmbFS.Text = $keepText
        } finally {
            $script:SuppressEvents = $false
        }
    }
}

#endregion


#region ---------- Active Directory ----------
# Операции с AD: параметры подключения, поиск групп и объектов, экранирование DN,
# создание OU с именем папки, создание групп, назначение NTFS-прав.

# Формирует параметры для команд AD: -Server и (при наличии) -Credential.
function Get-ADConnectionParams {
    $p = @{}
    if ($script:State.DC) { $p['Server'] = $script:State.DC }
    if ($script:State.Credential) { $p['Credential'] = $script:State.Credential }
    return $p
}

# Преобразует текст из поля DC (например, "dc01 (10.0.0.1)") в имя сервера.
function Resolve-DCFromText {
    param([string]$Text)
    $t = ($Text -replace '\s+', ' ').Trim()
    if (-not $t) { return '' }
    if ($script:DCMap -and $script:DCMap.ContainsKey($t)) { return $script:DCMap[$t] }
    if ($t -match '^([^\s(]+)\s*\(') { return $Matches[1] }
    return $t
}

# Ищет группу AD по имени; если группы нет - возвращает $null, а не ошибку.
function Find-ADGroupSafe {
    param([Parameter(Mandatory)][string]$Name)
    if (-not $script:State.ADReady) { throw 'Модуль ActiveDirectory недоступен.' }
    $p = Get-ADConnectionParams
    try {
        return (Get-ADGroup -Identity $Name @p -ErrorAction Stop)
    } catch {
        if ($_.FullyQualifiedErrorId -notmatch 'ADIdentityNotFoundException') { throw }
        $escaped = $Name.Replace("'", "''")
        $found = @(Get-ADGroup -Filter "Name -eq '$escaped'" @p -ErrorAction Stop)
        if ($found.Count -gt 0) { return $found[0] }
        return $null
    }
}

# Экранирует значение для DN (RDN): запятые, плюсы, решётки, пробелы и др.
function ConvertTo-DnRdnValue {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Value.Length; $i++) {
        $ch = $Value[$i]
        if ($ch -eq '\' -or $ch -eq ',' -or $ch -eq '+' -or $ch -eq '"' -or $ch -eq '<' -or $ch -eq '>' -or $ch -eq ';' -or $ch -eq '=') {
            [void]$sb.Append('\')
            [void]$sb.Append($ch)
        } else {
            [void]$sb.Append($ch)
        }
    }
    $escaped = $sb.ToString()
    if ($escaped.StartsWith('#')) { $escaped = '\' + $escaped }
    if ($escaped.StartsWith(' ')) { $escaped = '\' + $escaped }
    if ($escaped.EndsWith(' ')) { $escaped = $escaped.Substring(0, $escaped.Length - 1) + '\ ' }
    return $escaped
}

# Ищет объект AD по DN; если объекта нет - возвращает $null.
function Find-ADObjectSafe {
    param([Parameter(Mandatory)][string]$DistinguishedName)
    $p = Get-ADConnectionParams
    try {
        return (Get-ADObject -Identity $DistinguishedName @p -ErrorAction Stop)
    } catch {
        if ($_.FullyQualifiedErrorId -match 'ADIdentityNotFoundException') { return $null }
        throw
    }
}

# Возвращает DN OU с именем папки; создаёт её внутри выбранной OU, если её нет.
function Ensure-FolderOU {
    param(
        [Parameter(Mandatory)][string]$FolderName,
        [Parameter(Mandatory)][string]$BaseOU
    )
    $dn = 'OU=' + (ConvertTo-DnRdnValue -Value $FolderName) + ',' + $BaseOU
    # Если OU уже есть - используем её, иначе создаём.
    $existing = Find-ADObjectSafe -DistinguishedName $dn
    if ($existing) {
        $objClass = [string]$existing.ObjectClass
        if ($objClass -and $objClass -ne 'organizationalUnit' -and $objClass -ne 'container') {
            throw ('Объект с DN ' + $dn + ' уже существует, но это не OU и не контейнер (' + $objClass + '). Укажите другое подразделение.')
        }
        Write-UiLog ('OU уже существует: ' + $dn)
        return $dn
    }
    $p = Get-ADConnectionParams
    try {
        $null = New-ADOrganizationalUnit -Name $FolderName -Path $BaseOU @p -ErrorAction Stop
    } catch {
        throw ('Не удалось создать OU "' + $FolderName + '" в ' + $BaseOU + ': ' + $_.Exception.Message)
    }
    Write-UiLog ('Создана OU: ' + $dn) 'Ok'
    return $dn
}

# Проверяет, что указанная OU/контейнер существует и доступна.
function Assert-ADContainerExists {
    param([string]$DistinguishedName)
    $p = Get-ADConnectionParams
    try {
        $obj = Get-ADObject -Identity $DistinguishedName @p -ErrorAction Stop
        if (-not $obj) { throw 'объект не найден' }
    } catch {
        throw ('Подразделение (OU) не найдено или недоступно: ' + $DistinguishedName)
    }
}

# Описывает три группы: имена FS_<папка>_RO/RW/FA, права NTFS и описания.
function Get-GroupDefinitions {
    param([Parameter(Mandatory)][string]$FolderName)
    return @(
        [pscustomobject]@{
            Key         = 'RO'
            Name        = 'FS_' + $FolderName + '_RO'
            Rights      = 'ReadAndExecute'
            RightsText  = 'Только чтение'
            Description = 'Группа доступа к ' + $FolderName + ' (Только чтение)'
        }
        [pscustomobject]@{
            Key         = 'RW'
            Name        = 'FS_' + $FolderName + '_RW'
            Rights      = 'Modify'
            RightsText  = 'Чтение и Запись'
            Description = 'Группа доступа к ' + $FolderName + ' (Чтение и Запись)'
        }
        [pscustomobject]@{
            Key         = 'FA'
            Name        = 'FS_' + $FolderName + '_FA'
            Rights      = 'FullControl'
            RightsText  = 'Полный доступ'
            Description = 'Группа доступа к ' + $FolderName + ' (Полный доступ)'
        }
    )
}

# Находит существующие группы и создаёт отсутствующие в целевой OU.
function Ensure-FolderGroups {
    param(
        [Parameter(Mandatory)][string]$FolderName,
        [object[]]$Definitions,
        [string]$TargetOU = ''
    )
    if (-not $Definitions) { $Definitions = Get-GroupDefinitions -FolderName $FolderName }
    if (-not $TargetOU) { $TargetOU = $script:State.TargetOU }

    $p = Get-ADConnectionParams
    $entries = @()
    $toCreate = @()

    # Сначала ищем группы по именам; отсутствующие запоминаем для создания.
    foreach ($def in $Definitions) {
        $group = Find-ADGroupSafe -Name $def.Name
        if ($group) {
            $scope = [string]$group.GroupScope
            $category = [string]$group.GroupCategory
            Write-UiLog ('Группа уже существует: ' + $def.Name + ' (Scope: ' + $scope + ', Category: ' + $category + ')')
            if ($scope -ne 'DomainLocal' -or $category -ne 'Security') {
                Write-UiLog ('Внимание: тип группы ' + $def.Name + ' отличается от Domain Local / Security. Права будут назначены этой группе.') 'Warn'
            }
            $entries += [pscustomobject]@{ Definition = $def; Group = $group; Created = $false }
        } else {
            $toCreate += $def
        }
    }

    # Создаём только отсутствующие группы в целевой OU.
    if ($toCreate.Count -gt 0) {
        if (-not $TargetOU) { throw 'Не указано подразделение (OU), в котором создаются группы.' }
        Assert-ADContainerExists -DistinguishedName $TargetOU
        foreach ($def in $toCreate) {
            Write-UiLog ('Создание группы ' + $def.Name + ' в ' + $TargetOU + ' ...')
            $group = New-ADGroup -Name $def.Name -GroupScope DomainLocal -GroupCategory Security -Path $TargetOU -Description $def.Description -PassThru @p
            Write-UiLog ('Создана группа: ' + $def.Name + ' - ' + $def.Description) 'Ok'
            $entries += [pscustomobject]@{ Definition = $def; Group = $group; Created = $true }
        }
    }
    return $entries
}

# Назначает на папку ACE для трёх групп (идемпотентно, с наследованием).
function Set-FolderAcl {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][object[]]$GroupEntries
    )

    # Читаем текущий ACL папки (работает и по UNC-пути).
    $acl = Get-Acl -LiteralPath $FolderPath -ErrorAction Stop

    # Опционально отключаем наследование, сохраняя текущие права как явные.
    if ($script:State.DisableInheritance -and -not $acl.AreAccessRulesProtected) {
        $acl.SetAccessRuleProtection($true, $true)
        Write-UiLog 'Наследование разрешений отключено, текущие права скопированы как явные.' 'Warn'
    }

    # Для каждой группы убираем старые явные ACE и добавляем нужный.
    foreach ($entry in $GroupEntries) {
        $sidValue = ''
        try { $sidValue = [string]$entry.Group.SID.Value } catch { }
        if (-not $sidValue) {
            $fresh = Find-ADGroupSafe -Name $entry.Definition.Name
            if ($fresh -and $fresh.SID) { $sidValue = [string]$fresh.SID.Value }
        }
        if (-not $sidValue) { throw ('Не удалось получить SID группы ' + $entry.Definition.Name) }
        $sid = New-Object System.Security.Principal.SecurityIdentifier($sidValue)
        $rights = [System.Security.AccessControl.FileSystemRights]$entry.Definition.Rights
        $inherit = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit

        $remove = @()
        foreach ($rule in $acl.Access) {
            if ($rule.IsInherited) { continue }
            $ruleSid = $null
            try { $ruleSid = $rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]) } catch { }
            if ($ruleSid -and $ruleSid.Value -eq $sid.Value) { $remove += $rule }
        }
        foreach ($rule in $remove) {
            [void]$acl.RemoveAccessRuleSpecific($rule)
        }

        $ace = [System.Security.AccessControl.FileSystemAccessRule]::new(
            [System.Security.Principal.IdentityReference]$sid,
            $rights,
            $inherit,
            [System.Security.AccessControl.PropagationFlags]::None,
            [System.Security.AccessControl.AccessControlType]::Allow)
        $acl.AddAccessRule($ace)
        Write-UiLog ('Права "' + $entry.Definition.RightsText + '" (' + $entry.Definition.Rights + ') для ' + $entry.Definition.Name + ' назначены.')
    }

    # Записываем изменённый ACL обратно на файловый сервер.
    Set-Acl -LiteralPath $FolderPath -AclObject $acl -ErrorAction Stop
}

# Добавляет пользователей в выбранные группы, собирает ошибки.
function Add-UsersToFolderGroups {
    param(
        [Parameter(Mandatory)][string]$FolderName,
        [Parameter(Mandatory)][string[]]$SamAccounts,
        [Parameter(Mandatory)][string[]]$Keys
    )
    $errors = New-Object System.Collections.ArrayList
    $defs = @(Get-GroupDefinitions -FolderName $FolderName | Where-Object { $Keys -contains $_.Key })
    $p = Get-ADConnectionParams
    foreach ($def in $defs) {
        $group = Find-ADGroupSafe -Name $def.Name
        if (-not $group) {
            [void]$errors.Add('Группа ' + $def.Name + ' не найдена.')
            continue
        }
        foreach ($sam in $SamAccounts) {
            try {
                Add-ADGroupMember -Identity $group.DistinguishedName -Members $sam @p -ErrorAction Stop
                Write-UiLog ('В группу ' + $def.Name + ' добавлен: ' + $sam) 'Ok'
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match 'already a member|уже является членом') {
                    Write-UiLog ($sam + ' уже состоит в группе ' + $def.Name + '.')
                } else {
                    Write-UiLog ('Ошибка добавления ' + $sam + ' в ' + $def.Name + ': ' + $msg) 'Error'
                    [void]$errors.Add($sam + ' -> ' + $def.Name + ': ' + $msg)
                }
            }
        }
    }
    return @($errors)
}

#endregion


#region ---------- Основные операции ----------
# Домен и контроллеры, список папок, выбор папки, статусы групп, участники групп,
# а также главный сценарий обработки папки (папка -> OU -> группы -> права).

# Запрашивает список контроллеров домена и заполняет выпадающий список.
function Update-DCList {
    if (-not $script:State.ADReady) { throw 'Модуль ActiveDirectory недоступен.' }

    $dcs = @()
    $attempt = 0
    while ($true) {
        $p = @{}
        if ($script:State.Credential) { $p['Credential'] = $script:State.Credential }
        try {
            $dcs = @(Get-ADDomainController -Filter * @p -ErrorAction Stop | Sort-Object Name)
            break
        } catch {
            $msg = $_.Exception.Message
            if ($attempt -eq 0 -and (Test-LooksLikeAuthError $msg) -and (Request-ADCredentials -Reason $msg)) {
                $attempt++
                continue
            }
            throw
        }
    }

    Write-UiLog ('Найдено контроллеров домена: ' + $dcs.Count)

    $previous = ''
    if ($script:cmbDC) { $previous = $script:cmbDC.Text }
    $script:DCList = $dcs
    $script:DCMap = @{}

    $script:SuppressEvents = $true
    try {
        $script:cmbDC.BeginUpdate()
        try {
            $script:cmbDC.Items.Clear()
            foreach ($dc in $dcs) {
                $ip = ''
                $ips = @($dc.IPv4Address)
                if ($ips.Count -gt 0 -and $ips[0]) { $ip = [string]$ips[0] }
                $display = [string]$dc.Name
                if ($ip) { $display = $display + ' (' + $ip + ')' }
                [void]$script:cmbDC.Items.Add($display)
                $script:DCMap[$display] = [string]$dc.HostName
            }
        } finally {
            $script:cmbDC.EndUpdate()
        }
    } finally {
        $script:SuppressEvents = $false
    }

    $desired = $previous
    if (-not $desired) { $desired = $script:State.DC }
    $matched = ''
    if ($desired) {
        foreach ($key in $script:DCMap.Keys) {
            if ($key -like ($desired + '*') -or $script:DCMap[$key] -eq $desired) { $matched = $key; break }
        }
    }
    if (-not $matched -and $script:cmbDC.Items.Count -gt 0) { $matched = [string]$script:cmbDC.Items[0] }
    if ($matched) {
        $script:cmbDC.Text = $matched
        $script:State.DC = $script:DCMap[$matched]
    }
}

# Получает данные домена; при ошибке аутентификации предлагает ввести учётные данные.
function Initialize-DomainInfo {
    if (-not $script:State.ADReady) {
        $script:lblDomainInfo.Text = 'Домен: модуль ActiveDirectory недоступен'
        return
    }

    $domain = $null
    $attempt = 0
    while ($true) {
        $p = @{}
        if ($script:State.Credential) { $p['Credential'] = $script:State.Credential }
        try {
            $domain = Get-ADDomain @p -ErrorAction Stop
            break
        } catch {
            $msg = $_.Exception.Message
            Write-UiLog ('Не удалось получить данные домена: ' + $msg) 'Error'
            if ($attempt -eq 0 -and (Test-LooksLikeAuthError $msg) -and (Request-ADCredentials -Reason $msg)) {
                $attempt++
                continue
            }
            $script:lblDomainInfo.Text = 'Домен: ошибка'
            return
        }
    }

    $script:State.Domain = [string]$domain.DNSRoot
    $script:State.DomainDN = [string]$domain.DistinguishedName
    $script:State.NetBIOS = [string]$domain.NetBIOSName
    $script:lblDomainInfo.Text = 'Домен: ' + $domain.NetBIOSName + ' (' + $domain.DNSRoot + ')'
    Write-UiLog ('Домен: ' + $domain.DNSRoot + '; NetBIOS: ' + $domain.NetBIOSName + '; DN: ' + $domain.DistinguishedName)

    try {
        Update-DCList
    } catch {
        Write-UiLog ('Не удалось получить список контроллеров: ' + $_.Exception.Message) 'Error'
    }
}

# Показывает подпапки текущего пути (без выхода за корень).
function Update-FolderList {
    param([string]$SelectName)
    if (-not $script:lstFolders) { return }

    $script:SuppressEvents = $true
    $script:lstFolders.BeginUpdate()
    try {
        $script:lstFolders.Items.Clear()
        $script:State.SelectedPath = ''
        if ($script:lblSelected) { $script:lblSelected.Text = 'Выбранная папка: -' }
        if ($script:lvGroups) { $script:lvGroups.Items.Clear() }
        if ($script:lvMembers) { $script:lvMembers.Items.Clear() }
        if ($script:lblMembers) { $script:lblMembers.Text = 'Участники групп: -' }

        $path = $script:State.CurrentPath
        if (-not $path) { return }
        if (-not (Test-Path -LiteralPath $path -PathType Container)) {
            Write-UiLog ('Путь недоступен: ' + $path) 'Warn'
            return
        }
        $dirs = @(Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction Stop | Sort-Object Name)
        foreach ($d in $dirs) { [void]$script:lstFolders.Items.Add($d.Name) }
        Write-UiLog ('Папок в "' + $path + '": ' + $dirs.Count)

        if ($SelectName) {
            $idx = $script:lstFolders.Items.IndexOf($SelectName)
            if ($idx -ge 0) { $script:lstFolders.SelectedIndex = $idx }
        }
    } catch {
        Write-UiLog ('Не удалось прочитать список папок: ' + $_.Exception.Message) 'Error'
    } finally {
        $script:lstFolders.EndUpdate()
        $script:SuppressEvents = $false
    }
}

# Переходит в указанный каталог и обновляет список папок.
function Set-CurrentPath {
    param([Parameter(Mandatory)][string]$Path)
    $p = Normalize-Path $Path
    $script:State.CurrentPath = $p
    $script:txtCurrent.Text = $p
    $root = Normalize-Path $script:State.RootPath
    $script:btnUp.Enabled = ($p -ne $root)
    Update-FolderList
}

# Проверяет корневую папку и делает её текущей.
function Apply-RootPath {
    $root = Normalize-Path $script:txtRoot.Text
    if (-not $root) {
        Write-UiLog 'Не указана корневая папка.' 'Warn'
        return
    }
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        Write-UiLog ('Корневая папка недоступна: ' + $root) 'Error'
        Show-ErrorBox (Join-Lines @(
            'Папка не найдена или недоступна:',
            $root,
            '',
            'Проверьте путь (например \\server\share\папка), имя сервера и права доступа к папке.'))
        return
    }
    $script:State.RootPath = $root
    $script:State.FileServer = $script:cmbFS.Text.Trim()
    Add-ServerHistory $script:State.FileServer
    Write-UiLog ('Корневая папка: ' + $root)
    Set-CurrentPath $root
}

# Обрабатывает выбор папки: заголовки, статусы групп, участники.
function Set-Selection {
    param([Parameter(Mandatory)][string]$Path)
    $leaf = Get-PathLeaf -Path $Path
    $script:State.SelectedPath = $Path
    $script:lblSelected.Text = 'Выбранная папка: ' + $leaf
    $script:grpGroups.Text = 'Группы AD и права NTFS: ' + $leaf
    $script:grpUsers.Text = 'Пользователи в группах: ' + $leaf
    Update-GroupStatus -FolderPath $Path
    Update-MembersList -FolderPath $Path
}

# Заполняет таблицу: наличие групп в AD и назначенные ACE на папке.
function Update-GroupStatus {
    param([string]$FolderPath)
    if (-not $script:lvGroups) { return }

    $script:lvGroups.BeginUpdate()
    try {
        $script:lvGroups.Items.Clear()
        if (-not $FolderPath) { return }
        $leaf = Get-PathLeaf -Path $FolderPath
        $defs = Get-GroupDefinitions -FolderName $leaf

        $acl = $null
        try { $acl = Get-Acl -LiteralPath $FolderPath -ErrorAction Stop } catch { }

        foreach ($def in $defs) {
            $status = 'не проверено'
            $group = $null
            if ($script:State.ADReady) {
                try {
                    $group = Find-ADGroupSafe -Name $def.Name
                    if ($group) {
                        $status = 'есть (' + [string]$group.GroupScope + ' / ' + [string]$group.GroupCategory + ')'
                    } else {
                        $status = 'нет'
                    }
                } catch {
                    $status = 'ошибка проверки'
                }
            }

            $aceText = '-'
            $hasAce = $false
            if ($acl) {
                $aceText = 'нет'
                $groupSid = ''
                if ($group -and $group.SID) { $groupSid = [string]$group.SID.Value }
                foreach ($rule in $acl.Access) {
                    if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
                    $ruleSid = ''
                    try { $ruleSid = [string]$rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value } catch { }
                    if ($ruleSid -and $groupSid -and $ruleSid -eq $groupSid) { $hasAce = $true; break }
                    if (-not $groupSid -and $rule.IdentityReference.Value -like ('*' + $def.Name)) { $hasAce = $true; break }
                }
                if ($hasAce) { $aceText = 'назначены' }
            }

            $item = New-Object System.Windows.Forms.ListViewItem($def.Name)
            [void]$item.SubItems.Add($status)
            [void]$item.SubItems.Add($def.RightsText)
            [void]$item.SubItems.Add($aceText)
            $item.ToolTipText = $def.Description
            if (-not $script:State.ADReady -or -not $group) {
                $item.ForeColor = [System.Drawing.Color]::Gray
            } elseif (-not $hasAce) {
                $item.ForeColor = [System.Drawing.Color]::DarkOrange
            }
            [void]$script:lvGroups.Items.Add($item)
        }
    } catch {
        Write-UiLog ('Не удалось обновить статус групп: ' + $_.Exception.Message) 'Warn'
    } finally {
        $script:lvGroups.EndUpdate()
    }
}

# Показывает участников всех трёх групп выбранной папки.
function Update-MembersList {
    param([string]$FolderPath)
    if (-not $script:lvMembers) { return }

    $script:lvMembers.BeginUpdate()
    try {
        $script:lvMembers.Items.Clear()
        if (-not $FolderPath -or -not $script:State.ADReady) {
            $script:lblMembers.Text = 'Участники групп: -'
            return
        }
        $leaf = Get-PathLeaf -Path $FolderPath
        $defs = Get-GroupDefinitions -FolderName $leaf
        $p = Get-ADConnectionParams
        $total = 0

        foreach ($def in $defs) {
            $group = $null
            try { $group = Find-ADGroupSafe -Name $def.Name } catch { continue }
            if (-not $group) { continue }
            $members = @()
            try {
                $members = @(Get-ADGroupMember -Identity $group.DistinguishedName @p -ErrorAction Stop)
            } catch {
                Write-UiLog ('Не удалось прочитать участников ' + $def.Name + ': ' + $_.Exception.Message) 'Warn'
                continue
            }
            foreach ($m in $members) {
                $memberName = [string]$m.Name
                if (-not $memberName) { $memberName = [string]$m.SamAccountName }
                $sam = [string]$m.SamAccountName
                $cls = [string]$m.objectClass
                $item = New-Object System.Windows.Forms.ListViewItem($def.Name)
                [void]$item.SubItems.Add($memberName)
                [void]$item.SubItems.Add($sam)
                [void]$item.SubItems.Add($cls)
                $item.Tag = [pscustomobject]@{
                    GroupDN    = $group.DistinguishedName
                    GroupName  = $def.Name
                    MemberDN   = $m.DistinguishedName
                    MemberName = $memberName
                    Sam        = $sam
                }
                [void]$script:lvMembers.Items.Add($item)
                $total++
            }
        }
        $script:lblMembers.Text = 'Участники групп (всего: ' + $total + ')'
    } catch {
        Write-UiLog ('Не удалось прочитать участников групп: ' + $_.Exception.Message) 'Warn'
    } finally {
        $script:lvMembers.EndUpdate()
    }
}

# Главный сценарий: папка -> (OU) -> группы -> права NTFS.
function Invoke-FolderSetup {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [switch]$CreateFolder
    )

    $folderPath = Normalize-Path $FolderPath
    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Set-Status 'Обработка папки...'
    try {
        if (-not $script:State.ADReady) { throw 'Модуль ActiveDirectory недоступен - работа с группами невозможна.' }
        if (-not $script:State.DC) { throw 'Укажите контроллер домена.' }
        if (-not $script:State.RootPath) { throw 'Укажите корневую папку на файловом сервере.' }
        if (-not (Test-PathUnderRoot -Path $folderPath -Root $script:State.RootPath)) {
            throw ('Папка должна находиться внутри корневой папки: ' + $script:State.RootPath)
        }

        $leaf = Get-PathLeaf -Path $folderPath
        if (-not $leaf) { throw 'Не удалось определить имя папки.' }

        # Шаг 1. Смотрим, какие из трёх групп уже существуют в AD.
        $defs = @(Get-GroupDefinitions -FolderName $leaf)
        $missing = @()
        foreach ($d in $defs) {
            if (-not (Find-ADGroupSafe -Name $d.Name)) { $missing += $d }
        }
        $targetOU = $script:State.TargetOU
        # Шаг 2. Если групп не хватает - проверяем OU и допустимость имени папки.
        if ($missing.Count -gt 0) {
            Assert-GroupNameCompatibility -FolderName $leaf
            if (-not $targetOU) { throw 'Укажите подразделение (OU), в котором нужно создавать группы.' }
            Assert-ADContainerExists -DistinguishedName $targetOU
        }

        # Шаг 3. Создаём папку (или убеждаемся, что она существует).
        if ($CreateFolder) {
            if (Test-Path -LiteralPath $folderPath) { throw ('Папка уже существует: ' + $folderPath) }
            [void][System.IO.Directory]::CreateDirectory($folderPath)
            Write-UiLog ('Создана папка: ' + $folderPath) 'Ok'
        } elseif (-not (Test-Path -LiteralPath $folderPath -PathType Container)) {
            throw ('Папка не найдена: ' + $folderPath)
        }

        # Шаг 4. При включённой галочке создаём OU с именем папки в выбранной OU.
        if ($missing.Count -gt 0 -and $script:State.CreateFolderOU) {
            $targetOU = Ensure-FolderOU -FolderName $leaf -BaseOU $script:State.TargetOU
        }
        if ($targetOU -and $missing.Count -gt 0) { Write-UiLog ('Целевая OU для групп: ' + $targetOU) }

        # Шаг 5. Создаём отсутствующие группы в целевой OU и назначаем права на папку.
        $entries = Ensure-FolderGroups -FolderName $leaf -Definitions $defs -TargetOU $targetOU
        Set-FolderAcl -FolderPath $folderPath -GroupEntries $entries
        Write-UiLog ('Права RO/RW/FA назначены на папку: ' + $folderPath) 'Ok'

        # Шаг 6. Обновляем список папок и карточку выбранной папки.
        Update-FolderList -SelectName $leaf
        Set-Selection $folderPath
    } catch {
        Write-UiLog $_.Exception.Message 'Error'
        Show-ErrorBox $_.Exception.Message
    } finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
        Set-Status 'Готово'
    }
}

#endregion


#region ---------- Вспомогательные функции путей и AD ----------
# Корректная работа с UNC-путями (имя/родитель), разбор DN, дерево OU,
# распознавание ошибок аутентификации, запрос учётных данных, список шар сервера.

# Имя последнего элемента пути; корректно работает для \server\share.
function Get-PathLeaf {
    param([string]$Path)
    $p = Normalize-Path $Path
    if (-not $p) { return '' }
    if ($p -match '^[A-Za-z]:\\$') { return '' }
    if ($p -match '^\\\\[^\\]+$') { return $p.Substring(2) }
    $idx = $p.LastIndexOf('\')
    if ($idx -lt 0) { return $p }
    return $p.Substring($idx + 1)
}

# Родительский путь; корректно работает для \server\share.
function Get-ParentPath {
    param([string]$Path)
    $p = Normalize-Path $Path
    if (-not $p) { return '' }
    if ($p -match '^\\\\[^\\]+$') { return '' }
    if ($p -match '^[A-Za-z]:\\$') { return '' }
    $idx = $p.LastIndexOf('\')
    if ($idx -lt 0) { return '' }
    if ($idx -eq 2 -and $p.Length -gt 1 -and $p[1] -eq ':') { return $p.Substring(0, 3) }
    $parent = $p.Substring(0, $idx)
    if ($parent -match '^[A-Za-z]:$') { $parent = $parent + '\' }
    return $parent
}

# Разбирает DN на компоненты с учётом экранирования (\,).
function Split-DistinguishedName {
    param([Parameter(Mandatory)][string]$DistinguishedName)
    $parts = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $DistinguishedName.Length; $i++) {
        $ch = $DistinguishedName[$i]
        if ($ch -eq '\' -and ($i + 1) -lt $DistinguishedName.Length) {
            [void]$sb.Append($ch)
            [void]$sb.Append($DistinguishedName[$i + 1])
            $i++
            continue
        }
        if ($ch -eq ',') {
            $parts.Add($sb.ToString())
            [void]$sb.Clear()
            continue
        }
        [void]$sb.Append($ch)
    }
    $parts.Add($sb.ToString())
    return $parts.ToArray()
}

# Возвращает родительский DN.
function Get-DNParent {
    param([Parameter(Mandatory)][string]$DistinguishedName)
    $parts = @(Split-DistinguishedName -DistinguishedName $DistinguishedName)
    if ($parts.Count -le 1) { return '' }
    return (($parts[1..($parts.Count - 1)]) -join ',')
}

# Возвращает имя OU (значение первого RDN) из DN.
function Get-DNLeafName {
    param([Parameter(Mandatory)][string]$DistinguishedName)
    $parts = @(Split-DistinguishedName -DistinguishedName $DistinguishedName)
    if ($parts.Count -eq 0) { return '' }
    $rdn = [string]$parts[0]
    $idx = $rdn.IndexOf('=')
    if ($idx -lt 0) { return $rdn }
    $value = $rdn.Substring($idx + 1)
    $value = $value.Replace('\,', ',').Replace('\=', '=').Replace('\+', '+').Replace('\\', '\')
    return $value
}

# Рекурсивно ищет узел дерева по значению Tag (DN).
function Find-TreeNodeByTag {
    param($Node, [string]$Tag)
    if (-not $Node -or -not $Tag) { return $null }
    $nodeTag = ''
    if ($Node.Tag) { $nodeTag = [string]$Node.Tag }
    if ($nodeTag -and $nodeTag.Equals($Tag, [System.StringComparison]::OrdinalIgnoreCase)) { return $Node }
    foreach ($child in $Node.Nodes) {
        $found = Find-TreeNodeByTag -Node $child -Tag $Tag
        if ($found) { return $found }
    }
    return $null
}

# Определяет, похожа ли ошибка AD на проблему с учётными данными.
function Test-LooksLikeAuthError {
    param([string]$Message)
    if (-not $Message) { return $false }
    if ($Message -match 'rejected the client credentials') { return $true }
    if ($Message -match 'Authentication') { return $true }
    if ($Message -match 'Logon failure') { return $true }
    if ($Message -match 'access is denied') { return $true }
    if ($Message -match 'credential') { return $true }
    if ($Message -match 'подлинност') { return $true }
    if ($Message -match 'отказано в доступе') { return $true }
    if ($Message -match 'учётные данные') { return $true }
    if ($Message -match 'учетные данные') { return $true }
    return $false
}

# Запрашивает учётные данные для AD, если неявная аутентификация не работает.
function Request-ADCredentials {
    param([string]$Reason = '')
    if ($script:State.Credential) { return $false }
    if ($script:CredentialPromptDeclined) { return $false }

    $message = 'Не удалось выполнить запрос к Active Directory под текущей учётной записью.'
    if ($Reason) {
        $message = $message + [Environment]::NewLine + [Environment]::NewLine + 'Причина: ' + $Reason
    }
    $message = $message + [Environment]::NewLine + [Environment]::NewLine + 'Введите учётные данные для работы с AD (или нажмите "Отмена").'

    $userName = $env:USERNAME
    if ($env:USERDOMAIN) { $userName = $env:USERDOMAIN + '\' + $env:USERNAME }

    $cred = $null
    try {
        $cred = Get-Credential -UserName $userName -Message $message
    } catch {
        Write-UiLog ('Не удалось запросить учётные данные: ' + $_.Exception.Message) 'Warn'
        $cred = $null
    }
    if (-not $cred) {
        $script:CredentialPromptDeclined = $true
        Write-UiLog 'Запрос учётных данных для AD отменён. Можно отметить "Другая УЗ для AD" вручную.' 'Warn'
        return $false
    }

    $script:State.Credential = $cred
    if ($script:chkUseCred) {
        $script:SuppressEvents = $true
        try { $script:chkUseCred.Checked = $true } finally { $script:SuppressEvents = $false }
    }
    if ($script:lblCred) { $script:lblCred.Text = 'AD: ' + $cred.UserName }
    Write-UiLog ('AD-операции будут выполняться от имени ' + $cred.UserName + '.') 'Warn'
    return $true
}

# Получает список сетевых ресурсов сервера через оболочку Windows.
function Get-UNCShares {
    param([Parameter(Mandatory)][string]$ServerPath)
    $shares = @()
    $shell = $null
    try {
        $shell = New-Object -ComObject Shell.Application
        $ns = $shell.NameSpace($ServerPath)
        if (-not $ns) { throw ('Не удалось открыть ' + $ServerPath) }
        foreach ($item in $ns.Items()) {
            if ($item.IsFolder) {
                $name = [string]$item.Name
                if ($name) { $shares += $name }
            }
        }
    } finally {
        if ($shell) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch { }
        }
    }
    return @($shares)
}

#endregion


# Строит дерево OU (иерархия, фильтр, раскрытие узлов).
function Build-OUtree {
    param(
        [Parameter(Mandatory)]$Tree,
        [object[]]$Items,
        [string]$RootDN = '',
        [string]$RootText = '',
        [string]$Filter = '',
        [int]$ExpandThreshold = 120
    )
    $q = ([string]$Filter).Trim()
    $nodes = @($Items)
    if ($q) {
        $keep = @{}
        foreach ($item in @($Items | Where-Object { ([string]$_.Name).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 })) {
            $dn = [string]$item.DN
            while ($dn) {
                $keep[$dn.ToLower()] = $true
                $dn = Get-DNParent -DistinguishedName $dn
            }
        }
        if ($RootDN) { $keep[$RootDN.ToLower()] = $true }
        $nodes = @($Items | Where-Object { $keep.ContainsKey(([string]$_.DN).ToLower()) })
    }

    $Tree.BeginUpdate()
    try {
        $Tree.Nodes.Clear()
        $rootNode = New-Object System.Windows.Forms.TreeNode($RootText)
        $rootNode.Tag = $RootDN
        [void]$Tree.Nodes.Add($rootNode)
        $map = @{}
        if ($RootDN) { $map[$RootDN.ToLower()] = $rootNode }
        foreach ($item in ($nodes | Sort-Object Depth, Name)) {
            $parentNode = $null
            $key = ([string]$item.ParentDN).ToLower()
            if ($map.ContainsKey($key)) { $parentNode = $map[$key] }
            if (-not $parentNode) { $parentNode = $rootNode }
            $node = New-Object System.Windows.Forms.TreeNode([string]$item.Name)
            $node.Tag = [string]$item.DN
            [void]$parentNode.Nodes.Add($node)
            $map[([string]$item.DN).ToLower()] = $node
        }
        $rootNode.Expand()
        if ($q -or $nodes.Count -le $ExpandThreshold) { $rootNode.ExpandAll() }
    } finally {
        $Tree.EndUpdate()
    }
    return @{ Count = $nodes.Count; Total = @($Items).Count }
}

#region ---------- Диалоги ----------
# Модальные окна: ввод имени папки, обзор папок и сетевых ресурсов, дерево OU,
# выбор сервера из домена, поиск пользователей и добавление их в группы.

# Диалог ввода строки с проверкой (например, имя новой папки).
function Show-InputDialog {
    param(
        [string]$Title = 'Ввод значения',
        [string]$Prompt = '',
        [string]$DefaultValue = '',
        [scriptblock]$Validator
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.ClientSize = New-Object System.Drawing.Size(520, 190)
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $ctx = @{ Value = $null }

    $null = New-UILabel -Parent $dlg -Text $Prompt -X 14 -Y 12 -Width 490 -Height 40
    $txt = New-UITextBox -Parent $dlg -X 14 -Y 56 -Width 490
    $txt.Text = $DefaultValue
    $lblError = New-UILabel -Parent $dlg -Text '' -X 14 -Y 88 -Width 490 -Height 34
    $lblError.ForeColor = [System.Drawing.Color]::Firebrick

    $btnOk = New-UIButton -Parent $dlg -Text 'OK' -X 300 -Y 140 -Width 96 -Height 30
    $btnCancel = New-UIButton -Parent $dlg -Text 'Отмена' -X 404 -Y 140 -Width 100 -Height 30
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    $btnOk.Add_Click({
        $v = $txt.Text.Trim()
        if ($Validator) {
            $err = & $Validator $v
            if ($err) {
                $lblError.Text = [string]$err
                $txt.Focus()
                return
            }
        }
        $ctx.Value = $v
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })

    $dlg.Add_Shown({
        $txt.Focus()
        $txt.SelectAll()
    })

    $res = $dlg.ShowDialog($script:Form)
    $dlg.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $ctx.Value }
    return $null
}

# Обзор папок и сетевых ресурсов на файловом сервере.
function Show-FolderPickerDialog {
    param(
        [string]$InitialPath = '',
        [string]$Title = 'Выбор папки'
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.ClientSize = New-Object System.Drawing.Size(720, 520)
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $ctx = @{ Value = $null; ShareMode = $false }

    $null = New-UILabel -Parent $dlg -Text 'Путь:' -X 12 -Y 14 -Width 40
    $txt = New-UITextBox -Parent $dlg -X 52 -Y 12 -Width 540
    $txt.Text = $InitialPath
    $btnGo = New-UIButton -Parent $dlg -Text 'Перейти' -X 602 -Y 11 -Width 100

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(12, 44)
    $lst.Size = New-Object System.Drawing.Size(690, 370)
    $lst.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $lst.IntegralHeight = $false
    $lst.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $dlg.Controls.Add($lst)

    $null = New-UILabel -Parent $dlg -Text 'Совет: путь можно ввести вручную, например \\fileserver\Shared\Departments' -X 12 -Y 420 -Width 500
    $btnUp = New-UIButton -Parent $dlg -Text 'Вверх' -X 12 -Y 448 -Width 90
    $btnUp.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnOk = New-UIButton -Parent $dlg -Text 'Выбрать' -X 502 -Y 448 -Width 100
    $btnOk.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnCancel = New-UIButton -Parent $dlg -Text 'Отмена' -X 608 -Y 448 -Width 100
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $dlg.CancelButton = $btnCancel

    $navigate = {
        param([string]$p)
        $p = Normalize-Path $p
        if (-not $p) { return }

        if ($p -match '^\\\\[^\\]+$') {
            $shares = @()
            try { $shares = @(Get-UNCShares -ServerPath $p) } catch { $shares = @() }
            if ($shares.Count -eq 0) {
                [void][System.Windows.Forms.MessageBox]::Show($dlg,
                    ('Не удалось получить список сетевых ресурсов:' + [Environment]::NewLine + $p + [Environment]::NewLine + [Environment]::NewLine + 'Введите путь вручную в виде \\server\share\папка.'),
                    $dlg.Text, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $txt.Text = $p
                return
            }
            $txt.Text = $p
            $ctx.ShareMode = $true
            $lst.BeginUpdate()
            try {
                $lst.Items.Clear()
                foreach ($s in $shares) { [void]$lst.Items.Add($s) }
            } finally {
                $lst.EndUpdate()
            }
            return
        }

        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            [void][System.Windows.Forms.MessageBox]::Show($dlg,
                ('Путь не найден или недоступен:' + [Environment]::NewLine + $p),
                $dlg.Text, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $dirs = @()
        try { $dirs = @(Get-ChildItem -LiteralPath $p -Directory -Force -ErrorAction Stop | Sort-Object Name) } catch { $dirs = @() }

        $txt.Text = $p
        $ctx.ShareMode = $false
        $lst.BeginUpdate()
        try {
            $lst.Items.Clear()
            foreach ($d in $dirs) { [void]$lst.Items.Add($d.Name) }
        } finally {
            $lst.EndUpdate()
        }
    }

    $btnGo.Add_Click({ & $navigate $txt.Text })
    $txt.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { & $navigate $txt.Text }
    })
    $lst.Add_DoubleClick({
        if (-not $lst.SelectedItem) { return }
        $cur = Normalize-Path $txt.Text
        & $navigate (Join-Path $cur ([string]$lst.SelectedItem))
    })
    $btnUp.Add_Click({
        $parent = Get-ParentPath -Path $txt.Text
        if (-not $parent) { return }
        & $navigate $parent
    })
    $btnOk.Add_Click({
        $p = Normalize-Path $txt.Text
        if ($ctx.ShareMode -or ($p -match '^\\\\[^\\]+$')) {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, 'Сначала выберите сетевой ресурс (\\server\share), затем папку внутри него.', $dlg.Text,
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, 'Выбранный путь недоступен.', $dlg.Text,
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $ctx.Value = $p
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })

    $dlg.Add_Shown({
        if ($InitialPath) { & $navigate $InitialPath }
    })
    $res = $dlg.ShowDialog($script:Form)
    $dlg.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $ctx.Value }
    return $null
}

# Выбор OU в виде дерева с фильтром.
function Show-OUPickerDialog {
    param([string]$InitialValue = '')

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Выбор подразделения (OU)'
    $dlg.ClientSize = New-Object System.Drawing.Size(780, 600)
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $ctx = @{ All = @(); Result = $null; RootDN = ''; RootText = '' }

    $null = New-UILabel -Parent $dlg -Text 'Фильтр по имени OU:' -X 12 -Y 14 -Width 120
    $txtFilter = New-UITextBox -Parent $dlg -X 140 -Y 12 -Width 620
    $lblCount = New-UILabel -Parent $dlg -Text 'Загрузка...' -X 12 -Y 42 -Width 500

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Location = New-Object System.Drawing.Point(12, 66)
    $tree.Size = New-Object System.Drawing.Size(756, 450)
    $tree.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $tree.HideSelection = $false
    $tree.ShowLines = $true
    $tree.ShowRootLines = $true
    $tree.ShowPlusMinus = $true
    $dlg.Controls.Add($tree)

    $lblSelected = New-UILabel -Parent $dlg -Text 'Выбрано: (ничего)' -X 12 -Y 522 -Width 756
    $btnOk = New-UIButton -Parent $dlg -Text 'Выбрать' -X 562 -Y 552 -Width 100
    $btnOk.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnDomain = New-UIButton -Parent $dlg -Text 'Корень домена' -X 12 -Y 552 -Width 150
    $btnDomain.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnCancel = New-UIButton -Parent $dlg -Text 'Отмена' -X 668 -Y 552 -Width 100
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $dlg.CancelButton = $btnCancel

    $rebuild = {
        $stats = Build-OUtree -Tree $tree -Items $ctx.All -RootDN $ctx.RootDN -RootText $ctx.RootText -Filter $txtFilter.Text
        $lblCount.Text = 'Подразделений: ' + $stats.Count + ' из ' + $stats.Total
    }

    $dlg.Add_Shown({
        try {
            $p = Get-ADConnectionParams
            $ous = @(Get-ADOrganizationalUnit -Filter * @p -ErrorAction Stop)
            $list = @()
            foreach ($ou in $ous) {
                $dn = [string]$ou.DistinguishedName
                $parts = @(Split-DistinguishedName -DistinguishedName $dn)
                $parent = ''
                if ($parts.Count -gt 1) { $parent = (@($parts[1..($parts.Count - 1)]) -join ',') }
                $list += [pscustomobject]@{
                    DN       = $dn
                    Name     = (Get-DNLeafName -DistinguishedName $dn)
                    ParentDN = $parent
                    Depth    = $parts.Count - 1
                }
            }
            $ctx.All = @($list)
            $ctx.RootDN = $script:State.DomainDN
            if (-not $ctx.RootDN) {
                $first = $list | Select-Object -First 1
                if ($first) { $ctx.RootDN = $first.ParentDN }
            }
            $ctx.RootText = $script:State.Domain
            if (-not $ctx.RootText) { $ctx.RootText = $ctx.RootDN }
            & $rebuild
            if ($InitialValue -and $tree.Nodes.Count -gt 0) {
                $node = Find-TreeNodeByTag -Node $tree.Nodes[0] -Tag $InitialValue
                if ($node) {
                    $tree.SelectedNode = $node
                    $node.EnsureVisible()
                }
            }
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, ('Не удалось получить список OU:' + [Environment]::NewLine + $_.Exception.Message), 'Ошибка',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    $txtFilter.Add_TextChanged({ & $rebuild })
    $tree.Add_AfterSelect({
        if ($tree.SelectedNode) { $lblSelected.Text = 'Выбрано: ' + [string]$tree.SelectedNode.Tag }
    })
    $tree.Add_DoubleClick({
        if ($tree.SelectedNode -and $tree.SelectedNode.Tag) {
            $ctx.Result = [string]$tree.SelectedNode.Tag
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
    })
    $btnOk.Add_Click({
        if (-not $tree.SelectedNode) {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, 'Выберите подразделение в дереве.', $dlg.Text,
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $ctx.Result = [string]$tree.SelectedNode.Tag
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    $btnDomain.Add_Click({
        if (-not $ctx.RootDN) { return }
        $ctx.Result = $ctx.RootDN
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })

    $res = $dlg.ShowDialog($script:Form)
    $dlg.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $ctx.Result }
    return $null
}

# Выбор файлового сервера из компьютеров домена.
function Show-ComputerPickerDialog {
    param([string]$InitialFilter = '')

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Выбор файлового сервера (компьютеры домена)'
    $dlg.ClientSize = New-Object System.Drawing.Size(780, 540)
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $ctx = @{ Value = $null; All = @() }

    $null = New-UILabel -Parent $dlg -Text 'Фильтр по имени:' -X 12 -Y 14 -Width 110
    $txtFilter = New-UITextBox -Parent $dlg -X 126 -Y 12 -Width 420
    $txtFilter.Text = $InitialFilter
    $lblCount = New-UILabel -Parent $dlg -Text 'Загрузка...' -X 560 -Y 14 -Width 200

    $lst = New-Object System.Windows.Forms.ListView
    $lst.Location = New-Object System.Drawing.Point(12, 44)
    $lst.Size = New-Object System.Drawing.Size(756, 420)
    $lst.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $lst.View = 'Details'
    $lst.FullRowSelect = $true
    $lst.GridLines = $false
    $lst.HideSelection = $false
    $lst.MultiSelect = $false
    [void]$lst.Columns.Add('Имя', 180)
    [void]$lst.Columns.Add('DNS-имя', 260)
    [void]$lst.Columns.Add('Операционная система', 290)
    $dlg.Controls.Add($lst)

    $btnOk = New-UIButton -Parent $dlg -Text 'Выбрать' -X 548 -Y 476 -Width 105
    $btnOk.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnCancel = New-UIButton -Parent $dlg -Text 'Отмена' -X 663 -Y 476 -Width 105
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $dlg.CancelButton = $btnCancel
    $dlg.AcceptButton = $btnOk

    $applyFilter = {
        $q = $txtFilter.Text.Trim()
        $lst.BeginUpdate()
        try {
            $lst.Items.Clear()
            $shown = 0
            foreach ($c in $ctx.All) {
                if ($q) {
                    $match = ([string]$c.Name).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                    if (-not $match -and $c.DNSHostName) {
                        $match = ([string]$c.DNSHostName).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                    }
                    if (-not $match) { continue }
                }
                if ($shown -ge 500) { break }
                $dns = [string]$c.DNSHostName
                $os = [string]$c.OperatingSystem
                $item = New-Object System.Windows.Forms.ListViewItem([string]$c.Name)
                [void]$item.SubItems.Add($dns)
                [void]$item.SubItems.Add($os)
                if ($dns) { $item.Tag = $dns } else { $item.Tag = [string]$c.Name }
                [void]$lst.Items.Add($item)
                $shown++
            }
            $lblCount.Text = 'Показано: ' + $lst.Items.Count + ' из ' + $ctx.All.Count
        } finally {
            $lst.EndUpdate()
        }
    }

    $dlg.Add_Shown({
        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $p = Get-ADConnectionParams
            $computers = @(Get-ADComputer -Filter "OperatingSystem -like '*Server*'" -Properties DNSHostName, OperatingSystem -ResultSetSize 5000 @p -ErrorAction Stop | Sort-Object Name)
            $ctx.All = $computers
            & $applyFilter
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, ('Не удалось получить список компьютеров:' + [Environment]::NewLine + $_.Exception.Message), 'Ошибка',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $txtFilter.Add_TextChanged({ & $applyFilter })
    $btnOk.Add_Click({
        if (-not $lst.SelectedItems.Count) { return }
        $ctx.Value = [string]$lst.SelectedItems[0].Tag
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    $lst.Add_DoubleClick({
        if ($lst.SelectedItems.Count) {
            $ctx.Value = [string]$lst.SelectedItems[0].Tag
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
    })

    $res = $dlg.ShowDialog($script:Form)
    $dlg.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $ctx.Value }
    return $null
}

# Поиск пользователей AD и добавление в группы.
function Show-UserPickerDialog {
    param(
        [Parameter(Mandatory)][string]$FolderName,
        [Parameter(Mandatory)][object[]]$Definitions,
        [Parameter(Mandatory)][string[]]$AvailableKeys
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Добавление пользователей: ' + $FolderName
    $dlg.ClientSize = New-Object System.Drawing.Size(780, 590)
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $ctx = @{ Result = $null }

    $null = New-UILabel -Parent $dlg -Text 'Поиск пользователей (ФИО, логин, e-mail; не менее 2 символов):' -X 12 -Y 10 -Width 520
    $txtSearch = New-UITextBox -Parent $dlg -X 12 -Y 32 -Width 620
    $btnSearch = New-UIButton -Parent $dlg -Text 'Найти' -X 642 -Y 31 -Width 120
    $lblFound = New-UILabel -Parent $dlg -Text 'Найдено: 0' -X 12 -Y 60 -Width 400

    $lstUsers = New-Object System.Windows.Forms.ListView
    $lstUsers.Location = New-Object System.Drawing.Point(12, 82)
    $lstUsers.Size = New-Object System.Drawing.Size(750, 290)
    $lstUsers.View = 'Details'
    $lstUsers.FullRowSelect = $true
    $lstUsers.MultiSelect = $true
    $lstUsers.GridLines = $false
    $lstUsers.HideSelection = $false
    [void]$lstUsers.Columns.Add('ФИО', 300)
    [void]$lstUsers.Columns.Add('Логин', 160)
    [void]$lstUsers.Columns.Add('Включён', 70)
    [void]$lstUsers.Columns.Add('E-mail', 210)
    $dlg.Controls.Add($lstUsers)

    $gbGroups = New-Object System.Windows.Forms.GroupBox
    $gbGroups.Text = 'Добавить в группы'
    $gbGroups.Location = New-Object System.Drawing.Point(12, 380)
    $gbGroups.Size = New-Object System.Drawing.Size(750, 58)
    $dlg.Controls.Add($gbGroups)

    $chkRO = New-UICheckBox -Parent $gbGroups -Text ('RO: FS_' + $FolderName + '_RO') -X 10 -Y 22 -Width 240
    $chkRW = New-UICheckBox -Parent $gbGroups -Text ('RW: FS_' + $FolderName + '_RW') -X 254 -Y 22 -Width 240
    $chkFA = New-UICheckBox -Parent $gbGroups -Text ('FA: FS_' + $FolderName + '_FA') -X 498 -Y 22 -Width 240

    $checks = @(
        [pscustomobject]@{ Key = 'RO'; Box = $chkRO },
        [pscustomobject]@{ Key = 'RW'; Box = $chkRW },
        [pscustomobject]@{ Key = 'FA'; Box = $chkFA }
    )
    foreach ($c in $checks) {
        if ($AvailableKeys -contains $c.Key) {
            if ($c.Key -eq 'RW') { $c.Box.Checked = $true }
        } else {
            $c.Box.Enabled = $false
            $c.Box.Text = $c.Box.Text + ' (нет в AD)'
        }
    }

    $null = New-UILabel -Parent $dlg -Text 'Дополнительно: логины через ";" (если поиск не находит пользователя):' -X 12 -Y 446 -Width 560
    $txtManual = New-UITextBox -Parent $dlg -X 12 -Y 466 -Width 750

    $btnAdd = New-UIButton -Parent $dlg -Text 'Добавить' -X 532 -Y 516 -Width 110 -Height 30
    $btnCancel = New-UIButton -Parent $dlg -Text 'Отмена' -X 652 -Y 516 -Width 110 -Height 30
    $dlg.AcceptButton = $btnSearch
    $dlg.CancelButton = $btnCancel

    $doSearch = {
        $q = $txtSearch.Text.Trim()
        if ($q.Length -lt 2) {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, 'Введите не менее 2 символов для поиска.', 'Поиск',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $p = Get-ADConnectionParams
            $escaped = $q.Replace("'", "''")
            $filter = "SamAccountName -like '*" + $escaped + "*' -or DisplayName -like '*" + $escaped + "*' -or Name -like '*" + $escaped + "*' -or Mail -like '*" + $escaped + "*'"
            $users = @(Get-ADUser -Filter $filter -Properties DisplayName, Mail, Enabled, Title -ResultSetSize 300 @p -ErrorAction Stop | Sort-Object DisplayName, Name)
            $lstUsers.BeginUpdate()
            try {
                $lstUsers.Items.Clear()
                foreach ($u in $users) {
                    $display = [string]$u.DisplayName
                    if (-not $display) { $display = [string]$u.Name }
                    $item = New-Object System.Windows.Forms.ListViewItem($display)
                    [void]$item.SubItems.Add([string]$u.SamAccountName)
                    $enabledText = 'нет'
                    if ($u.Enabled) { $enabledText = 'да' }
                    [void]$item.SubItems.Add($enabledText)
                    [void]$item.SubItems.Add([string]$u.Mail)
                    $item.Tag = [string]$u.SamAccountName
                    if (-not $u.Enabled) { $item.ForeColor = [System.Drawing.Color]::Gray }
                    [void]$lstUsers.Items.Add($item)
                }
            } finally {
                $lstUsers.EndUpdate()
            }
            $lblFound.Text = 'Найдено: ' + $users.Count + ' (показано не более 300)'
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, ('Ошибка поиска:' + [Environment]::NewLine + $_.Exception.Message), 'Поиск',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnSearch.Add_Click({ & $doSearch })

    $btnAdd.Add_Click({
        $sams = New-Object System.Collections.ArrayList
        foreach ($item in $lstUsers.SelectedItems) { [void]$sams.Add([string]$item.Tag) }
        $manual = $txtManual.Text.Trim()
        if ($manual) {
            foreach ($part in ($manual -split '[;,\s]+')) {
                if ($part) { [void]$sams.Add($part) }
            }
        }
        $unique = @($sams | Select-Object -Unique)
        if ($unique.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, 'Выберите пользователей в списке или введите логины вручную.', 'Добавление',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $keys = @()
        foreach ($c in $checks) {
            if ($c.Box.Checked) { $keys += $c.Key }
        }
        if ($keys.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($dlg, 'Отметьте хотя бы одну группу.', 'Добавление',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $ctx.Result = [pscustomobject]@{ SamAccounts = $unique; Keys = $keys }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })

    $res = $dlg.ShowDialog($script:Form)
    $dlg.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $ctx.Result }
    return $null
}

#endregion


#region ---------- Построение главного окна ----------
# Компоновка окна (панели, таблицы, журнал) и подписки на события.

# Создаёт главное окно, все панели, таблицы и подписки на события.
function Build-UI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $script:AppTitle
    $form.ClientSize = New-Object System.Drawing.Size(1200, 880)
    $form.MinimumSize = New-Object System.Drawing.Size(1180, 860)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $script:Form = $form

    $rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $rootLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rootLayout.ColumnCount = 1
    $rootLayout.RowCount = 4
    $rootLayout.Padding = New-Object System.Windows.Forms.Padding(6)
    [void]$rootLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$rootLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 208))
    [void]$rootLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$rootLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 168))
    [void]$rootLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 24))
    $form.Controls.Add($rootLayout)

    # ---------- Подключение ----------
    # Параметры: контроллер домена, файловый сервер, корневая папка,
    # подразделение (OU) и (опционально) отдельные учётные данные для AD.
    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = 'Подключение'
    $grpConn.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpConn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)

    $null = New-UILabel -Parent $grpConn -Text 'Контроллер домена:' -X 12 -Y 25 -Width 130
    $script:cmbDC = New-UIComboBox -Parent $grpConn -X 148 -Y 22 -Width 420 -Editable
    $btnRefreshDC = New-UIButton -Parent $grpConn -Text 'Обновить' -X 576 -Y 21 -Width 100
    $script:lblDomainInfo = New-UILabel -Parent $grpConn -Text 'Домен: -' -X 690 -Y 25 -Width 445

    $null = New-UILabel -Parent $grpConn -Text 'Файловый сервер:' -X 12 -Y 57 -Width 130
    $script:cmbFS = New-UIComboBox -Parent $grpConn -X 148 -Y 54 -Width 420 -Editable
    $btnServerList = New-UIButton -Parent $grpConn -Text 'Из домена...' -X 576 -Y 53 -Width 100
    $script:chkUseCred = New-UICheckBox -Parent $grpConn -Text 'Другая УЗ для AD' -X 690 -Y 55 -Width 165
    $script:lblCred = New-UILabel -Parent $grpConn -Text 'AD: текущая УЗ' -X 858 -Y 58 -Width 270

    $null = New-UILabel -Parent $grpConn -Text 'Корневая папка:' -X 12 -Y 89 -Width 130
    $script:txtRoot = New-UITextBox -Parent $grpConn -X 148 -Y 86 -Width 420
    $btnBrowseRoot = New-UIButton -Parent $grpConn -Text 'Обзор...' -X 576 -Y 85 -Width 100
    $null = New-UILabel -Parent $grpConn -Text 'Пример: \\fileserver\Shared\Departments' -X 690 -Y 89 -Width 445

    $null = New-UILabel -Parent $grpConn -Text 'Подразделение (OU):' -X 12 -Y 121 -Width 130
    $script:txtOU = New-UITextBox -Parent $grpConn -X 148 -Y 118 -Width 420
    $btnPickOU = New-UIButton -Parent $grpConn -Text 'Выбрать...' -X 576 -Y 117 -Width 100
    $null = New-UILabel -Parent $grpConn -Text 'Где создаются группы: OU=FileShares,DC=example,DC=local' -X 690 -Y 121 -Width 445
    $script:chkCreateFolderOU = New-UICheckBox -Parent $grpConn -Text 'Создавать OU с именем папки' -X 148 -Y 150 -Width 330
    $script:lblCreateOuHint = New-UILabel -Parent $grpConn -Text 'группы создаются в OU=ИмяПапки внутри выбранной OU' -X 486 -Y 153 -Width 660

    $rootLayout.Controls.Add($grpConn, 0, 0)

    # ---------- Тело ----------
    # Слева - работа с папками, справа - группы/права и пользователи.
    $bodyLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $bodyLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $bodyLayout.ColumnCount = 2
    $bodyLayout.RowCount = 1
    $bodyLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    [void]$bodyLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 43))
    [void]$bodyLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 57))
    [void]$bodyLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    $rootLayout.Controls.Add($bodyLayout, 0, 1)

    # ---------- Папки ----------
    # Текущий путь, список подпапок, навигация и создание папки.
    $grpFolders = New-Object System.Windows.Forms.GroupBox
    $grpFolders.Text = 'Папка на файловом сервере'
    $grpFolders.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpFolders.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)

    $tlpFolders = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpFolders.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tlpFolders.ColumnCount = 1
    $tlpFolders.RowCount = 3
    $tlpFolders.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 6)
    [void]$tlpFolders.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$tlpFolders.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 48))
    [void]$tlpFolders.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$tlpFolders.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 76))
    $grpFolders.Controls.Add($tlpFolders)

    $pnlPath = New-Object System.Windows.Forms.Panel
    $pnlPath.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlPath.Margin = New-Object System.Windows.Forms.Padding(0)
    $null = New-UILabel -Parent $pnlPath -Text 'Текущий путь:' -X 0 -Y 0 -Width 200
    $script:txtCurrent = New-UITextBox -Parent $pnlPath -X 0 -Y 20 -Width 470 -ReadOnly
    $tlpFolders.Controls.Add($pnlPath, 0, 0)

    $script:lstFolders = New-Object System.Windows.Forms.ListBox
    $script:lstFolders.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:lstFolders.IntegralHeight = $false
    $script:lstFolders.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $script:lstFolders.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
    $tlpFolders.Controls.Add($script:lstFolders, 0, 1)

    $pnlFolderButtons = New-Object System.Windows.Forms.Panel
    $pnlFolderButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlFolderButtons.Margin = New-Object System.Windows.Forms.Padding(0)
    $script:lblSelected = New-UILabel -Parent $pnlFolderButtons -Text 'Выбранная папка: -' -X 0 -Y 2 -Width 470
    $script:btnRefreshFolders = New-UIButton -Parent $pnlFolderButtons -Text 'Обновить' -X 0 -Y 28 -Width 95
    $script:btnUp = New-UIButton -Parent $pnlFolderButtons -Text 'Вверх' -X 101 -Y 28 -Width 75
    $script:btnCreateFolder = New-UIButton -Parent $pnlFolderButtons -Text 'Создать папку...' -X 182 -Y 28 -Width 145
    $tlpFolders.Controls.Add($pnlFolderButtons, 0, 2)

    $bodyLayout.Controls.Add($grpFolders, 0, 0)

    # ---------- Права и группы + пользователи ----------
    # Статусы трёх групп, назначение прав и работа с участниками.
    $rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $rightLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightLayout.ColumnCount = 1
    $rightLayout.RowCount = 2
    $rightLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    [void]$rightLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$rightLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 60))
    [void]$rightLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 40))
    $bodyLayout.Controls.Add($rightLayout, 1, 0)

    $script:grpGroups = New-Object System.Windows.Forms.GroupBox
    $script:grpGroups.Text = 'Группы AD и права NTFS'
    $script:grpGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:grpGroups.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)

    $tlpGroups = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tlpGroups.ColumnCount = 1
    $tlpGroups.RowCount = 2
    [void]$tlpGroups.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$tlpGroups.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$tlpGroups.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 92))
    $script:grpGroups.Controls.Add($tlpGroups)

    $script:lvGroups = New-Object System.Windows.Forms.ListView
    $script:lvGroups.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:lvGroups.View = 'Details'
    $script:lvGroups.FullRowSelect = $true
    $script:lvGroups.GridLines = $true
    $script:lvGroups.HideSelection = $false
    $script:lvGroups.ShowItemToolTips = $true
    $script:lvGroups.Margin = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)
    [void]$script:lvGroups.Columns.Add('Группа', 200)
    [void]$script:lvGroups.Columns.Add('Статус в AD', 150)
    [void]$script:lvGroups.Columns.Add('Права NTFS', 110)
    [void]$script:lvGroups.Columns.Add('ACE на папке', 110)
    $tlpGroups.Controls.Add($script:lvGroups, 0, 0)

    $pnlGroupButtons = New-Object System.Windows.Forms.Panel
    $pnlGroupButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlGroupButtons.Margin = New-Object System.Windows.Forms.Padding(0)
    $script:chkInherit = New-UICheckBox -Parent $pnlGroupButtons -Text 'Отключить наследование разрешений папки (текущие права копируются)' -X 6 -Y 2 -Width 600
    $script:btnApply = New-UIButton -Parent $pnlGroupButtons -Text 'Создать группы и назначить права' -X 6 -Y 28 -Width 305 -Height 30
    $script:btnRefreshStatus = New-UIButton -Parent $pnlGroupButtons -Text 'Обновить статус' -X 317 -Y 28 -Width 150 -Height 30
    $tlpGroups.Controls.Add($pnlGroupButtons, 0, 1)

    $rightLayout.Controls.Add($script:grpGroups, 0, 0)

    $script:grpUsers = New-Object System.Windows.Forms.GroupBox
    $script:grpUsers.Text = 'Пользователи в группах'
    $script:grpUsers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:grpUsers.Margin = New-Object System.Windows.Forms.Padding(0)

    $tlpUsers = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpUsers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tlpUsers.ColumnCount = 1
    $tlpUsers.RowCount = 3
    [void]$tlpUsers.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$tlpUsers.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 24))
    [void]$tlpUsers.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 38))
    [void]$tlpUsers.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    $script:grpUsers.Controls.Add($tlpUsers)

    $script:lblMembers = New-UILabel -Parent $script:grpUsers -Text 'Участники групп: -' -X 0 -Y 0 -Width 400
    $script:lblMembers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:lblMembers.Margin = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)
    $tlpUsers.Controls.Add($script:lblMembers, 0, 0)

    $pnlUserButtons = New-Object System.Windows.Forms.Panel
    $pnlUserButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlUserButtons.Margin = New-Object System.Windows.Forms.Padding(0)
    $script:btnAddUsers = New-UIButton -Parent $pnlUserButtons -Text 'Добавить пользователей...' -X 6 -Y 2 -Width 215
    $script:btnRemoveMember = New-UIButton -Parent $pnlUserButtons -Text 'Удалить выбранных' -X 227 -Y 2 -Width 165
    $script:btnRefreshMembers = New-UIButton -Parent $pnlUserButtons -Text 'Обновить' -X 398 -Y 2 -Width 110
    $tlpUsers.Controls.Add($pnlUserButtons, 0, 1)

    $script:lvMembers = New-Object System.Windows.Forms.ListView
    $script:lvMembers.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:lvMembers.View = 'Details'
    $script:lvMembers.FullRowSelect = $true
    $script:lvMembers.GridLines = $true
    $script:lvMembers.HideSelection = $false
    $script:lvMembers.MultiSelect = $true
    $script:lvMembers.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 4)
    [void]$script:lvMembers.Columns.Add('Группа', 170)
    [void]$script:lvMembers.Columns.Add('Пользователь', 200)
    [void]$script:lvMembers.Columns.Add('Логин', 140)
    [void]$script:lvMembers.Columns.Add('Тип', 70)
    $tlpUsers.Controls.Add($script:lvMembers, 0, 2)

    $rightLayout.Controls.Add($script:grpUsers, 0, 1)

    # ---------- Журнал ----------
    # Все действия и ошибки; очистка, копирование и сохранение журнала.
    $grpLog = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text = 'Журнал'
    $grpLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpLog.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)

    $tlpLog = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tlpLog.ColumnCount = 1
    $tlpLog.RowCount = 2
    [void]$tlpLog.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$tlpLog.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
    [void]$tlpLog.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    $grpLog.Controls.Add($tlpLog)

    $pnlLogButtons = New-Object System.Windows.Forms.Panel
    $pnlLogButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlLogButtons.Margin = New-Object System.Windows.Forms.Padding(0)
    $btnClearLog = New-UIButton -Parent $pnlLogButtons -Text 'Очистить' -X 6 -Y 2 -Width 95 -Height 25
    $btnSaveLog = New-UIButton -Parent $pnlLogButtons -Text 'Сохранить в файл...' -X 107 -Y 2 -Width 150 -Height 25
    $btnCopyLog = New-UIButton -Parent $pnlLogButtons -Text 'Копировать' -X 263 -Y 2 -Width 110 -Height 25
    $tlpLog.Controls.Add($pnlLogButtons, 0, 0)

    $script:LogBox = New-Object System.Windows.Forms.RichTextBox
    $script:LogBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:LogBox.ReadOnly = $true
    $script:LogBox.BackColor = [System.Drawing.Color]::White
    $script:LogBox.Font = New-Object System.Drawing.Font('Consolas', 9)
    $script:LogBox.WordWrap = $true
    $script:LogBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $script:LogBox.DetectUrls = $false
    $script:LogBox.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 4)
    $tlpLog.Controls.Add($script:LogBox, 0, 1)

    $rootLayout.Controls.Add($grpLog, 0, 2)

    # ---------- Строка состояния ----------
    # Короткий текст о текущей операции.
    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Text = 'Готово'
    $script:StatusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:StatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $script:StatusLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $script:StatusLabel.Padding = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $rootLayout.Controls.Add($script:StatusLabel, 0, 3)

    # ---------- Обработчики ----------
    # Ниже подключаются события элементов окна. Обработчики используют
    # переменные и функции скрипта, поэтому доступны после создания окна.
    $script:cmbDC.Add_TextChanged({
        if ($script:SuppressEvents) { return }
        $script:State.DC = Resolve-DCFromText $script:cmbDC.Text
    })

    $script:cmbDC.Add_SelectedIndexChanged({
        if ($script:SuppressEvents) { return }
        $script:State.DC = Resolve-DCFromText $script:cmbDC.Text
    })

    $btnRefreshDC.Add_Click({
        try {
            Set-Status 'Запрос списка контроллеров домена...'
            Update-DCList
            Write-UiLog 'Список контроллеров домена обновлён.' 'Ok'
        } catch {
            Write-UiLog ('Не удалось получить список контроллеров: ' + $_.Exception.Message) 'Error'
            Show-ErrorBox ('Не удалось получить список контроллеров домена:' + [Environment]::NewLine + $_.Exception.Message)
        } finally {
            Set-Status 'Готово'
        }
    })

    $script:cmbFS.Add_TextChanged({
        if ($script:SuppressEvents) { return }
        $script:State.FileServer = $script:cmbFS.Text.Trim()
    })

    $btnServerList.Add_Click({
        try {
            $sel = Show-ComputerPickerDialog -InitialFilter $script:cmbFS.Text
            if ($sel) {
                $script:cmbFS.Text = $sel
                $script:State.FileServer = $sel
                Add-ServerHistory $sel
            }
        } catch {
            Write-UiLog $_.Exception.Message 'Error'
            Show-ErrorBox $_.Exception.Message
        }
    })

    $btnBrowseRoot.Add_Click({
        try {
            $start = $script:txtRoot.Text.Trim()
            if (-not $start) {
                $fs = $script:cmbFS.Text.Trim()
                if ($fs) { $start = '\\' + $fs }
            }
            $sel = Show-FolderPickerDialog -InitialPath $start -Title 'Выбор корневой папки на файловом сервере'
            if ($sel) {
                $script:txtRoot.Text = $sel
                Apply-RootPath
            }
        } catch {
            Write-UiLog $_.Exception.Message 'Error'
            Show-ErrorBox $_.Exception.Message
        }
    })

    $btnPickOU.Add_Click({
        try {
            $sel = Show-OUPickerDialog -InitialValue $script:txtOU.Text.Trim()
            if ($sel) {
                $script:txtOU.Text = $sel
                $script:State.TargetOU = $sel
                Write-UiLog ('Подразделение для групп: ' + $sel)
            }
        } catch {
            Write-UiLog $_.Exception.Message 'Error'
            Show-ErrorBox $_.Exception.Message
        }
    })

    $script:txtOU.Add_TextChanged({
        if ($script:SuppressEvents) { return }
        $script:State.TargetOU = $script:txtOU.Text.Trim()
    })

    $script:txtRoot.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            try { Apply-RootPath } catch {
                Write-UiLog $_.Exception.Message 'Error'
                Show-ErrorBox $_.Exception.Message
            }
        }
    })

    $script:chkUseCred.Add_CheckedChanged({
        if ($script:SuppressEvents) { return }
        if ($script:chkUseCred.Checked) {
            $cred = Get-Credential -Message 'Учётные данные для операций с Active Directory'
            if (-not $cred) {
                $script:SuppressEvents = $true
                $script:chkUseCred.Checked = $false
                $script:SuppressEvents = $false
                return
            }
            $script:State.Credential = $cred
            $script:lblCred.Text = 'AD: ' + $cred.UserName
            Write-UiLog ('AD-операции будут выполняться от имени ' + $cred.UserName + '.') 'Warn'
            try {
                Initialize-DomainInfo
            } catch {
                Write-UiLog ('Не удалось обновить данные домена: ' + $_.Exception.Message) 'Error'
            }
        } else {
            $script:State.Credential = $null
            $script:CredentialPromptDeclined = $false
            $script:lblCred.Text = 'AD: текущая УЗ'
            Write-UiLog 'AD-операции выполняются от текущей учётной записи.'
        }
    })

    $script:btnRefreshFolders.Add_Click({
        try {
            $r = Normalize-Path $script:txtRoot.Text
            if ($r -ne $script:State.RootPath) {
                Apply-RootPath
            } else {
                Update-FolderList
            }
        } catch {
            Write-UiLog $_.Exception.Message 'Error'
            Show-ErrorBox $_.Exception.Message
        }
    })

    $script:btnUp.Add_Click({
        $root = Normalize-Path $script:State.RootPath
        $cur = Normalize-Path $script:State.CurrentPath
        if (-not $cur -or $cur -eq $root) { return }
        $parent = Get-ParentPath -Path $cur
        if (-not $parent) { $parent = $root }
        if (-not (Test-PathUnderRoot -Path $parent -Root $root)) { $parent = $root }
        Set-CurrentPath $parent
    })

    $script:lstFolders.Add_SelectedIndexChanged({
        if ($script:SuppressEvents) { return }
        try {
            if ($script:lstFolders.SelectedItem) {
                $path = Join-Path $script:State.CurrentPath ([string]$script:lstFolders.SelectedItem)
                Set-Selection $path
            }
        } catch {
            Write-UiLog ('Ошибка при выборе папки: ' + $_.Exception.Message) 'Error'
        }
    })

    $script:lstFolders.Add_DoubleClick({
        if ($script:SuppressEvents) { return }
        if (-not $script:lstFolders.SelectedItem) { return }
        try {
            $child = Join-Path $script:State.CurrentPath ([string]$script:lstFolders.SelectedItem)
            if (Test-Path -LiteralPath $child -PathType Container) {
                Set-CurrentPath $child
            }
        } catch {
            Write-UiLog ('Ошибка перехода в папку: ' + $_.Exception.Message) 'Error'
        }
    })

    $script:btnCreateFolder.Add_Click({
        try {
            if (-not $script:State.CurrentPath) {
                Show-WarnBox 'Сначала выберите корневую папку на файловом сервере.'
                return
            }
            $validator = {
                param($name)
                Test-NewFolderName -Name $name -ParentPath $script:State.CurrentPath
            }
            $name = Show-InputDialog -Title 'Создание папки' -Prompt ('Новая папка в:' + [Environment]::NewLine + $script:State.CurrentPath) -DefaultValue '' -Validator $validator
            if (-not $name) { return }
            $newPath = Join-Path $script:State.CurrentPath $name
            Invoke-FolderSetup -FolderPath $newPath -CreateFolder
        } catch {
            Write-UiLog $_.Exception.Message 'Error'
            Show-ErrorBox $_.Exception.Message
        }
    })

    $script:btnApply.Add_Click({
        if (-not $script:State.SelectedPath) {
            Show-WarnBox 'Выберите папку в списке слева.'
            return
        }
        Invoke-FolderSetup -FolderPath $script:State.SelectedPath
    })

    $script:btnRefreshStatus.Add_Click({
        if ($script:State.SelectedPath) {
            Update-GroupStatus -FolderPath $script:State.SelectedPath
            Update-MembersList -FolderPath $script:State.SelectedPath
        }
    })

    $script:chkInherit.Add_CheckedChanged({
        $script:State.DisableInheritance = [bool]$script:chkInherit.Checked
    })

    $script:chkCreateFolderOU.Add_CheckedChanged({
        if ($script:SuppressEvents) { return }
        $script:State.CreateFolderOU = [bool]$script:chkCreateFolderOU.Checked
        Update-CreateOuHint
    })

    $script:btnAddUsers.Add_Click({
        try {
            if (-not $script:State.SelectedPath) {
                Show-WarnBox 'Выберите папку в списке слева.'
                return
            }
            $leaf = Get-PathLeaf -Path $script:State.SelectedPath
            $defs = @(Get-GroupDefinitions -FolderName $leaf)
            $existing = @()
            foreach ($d in $defs) {
                if (Find-ADGroupSafe -Name $d.Name) { $existing += $d.Key }
            }
            if ($existing.Count -eq 0) {
                Show-WarnBox ('Группы для папки "' + $leaf + '" ещё не созданы.' + [Environment]::NewLine + 'Нажмите "Создать группы и назначить права".')
                return
            }
            $result = Show-UserPickerDialog -FolderName $leaf -Definitions $defs -AvailableKeys $existing
            if (-not $result) { return }
            $errors = @(Add-UsersToFolderGroups -FolderName $leaf -SamAccounts $result.SamAccounts -Keys $result.Keys)
            if ($errors.Count -gt 0) {
                Show-WarnBox ('Не удалось добавить некоторых пользователей:' + [Environment]::NewLine + ($errors -join [Environment]::NewLine))
            }
            Update-MembersList -FolderPath $script:State.SelectedPath
        } catch {
            Write-UiLog $_.Exception.Message 'Error'
            Show-ErrorBox $_.Exception.Message
        }
    })

    $script:btnRemoveMember.Add_Click({
        $items = @($script:lvMembers.SelectedItems)
        if ($items.Count -eq 0) {
            Show-WarnBox 'Выберите участников в списке.'
            return
        }
        $names = @()
        foreach ($item in $items) { $names += ($item.Tag.Sam + ' [' + $item.Tag.GroupName + ']') }
        if (-not (Confirm-Action ('Удалить из групп:' + [Environment]::NewLine + ($names -join [Environment]::NewLine)))) { return }
        $p = Get-ADConnectionParams
        foreach ($item in $items) {
            $t = $item.Tag
            try {
                Remove-ADGroupMember -Identity $t.GroupDN -Members $t.MemberDN -Confirm:$false @p -ErrorAction Stop
                Write-UiLog ('Удалён из группы ' + $t.GroupName + ': ' + $t.Sam) 'Ok'
            } catch {
                Write-UiLog ('Ошибка удаления ' + $t.Sam + ' из ' + $t.GroupName + ': ' + $_.Exception.Message) 'Error'
            }
        }
        Update-MembersList -FolderPath $script:State.SelectedPath
    })

    $script:btnRefreshMembers.Add_Click({
        Update-MembersList -FolderPath $script:State.SelectedPath
    })

    $btnClearLog.Add_Click({ $script:LogBox.Clear() })

    $btnCopyLog.Add_Click({
        try {
            if ($script:LogBox.TextLength -gt 0) {
                [System.Windows.Forms.Clipboard]::SetText($script:LogBox.Text)
                Set-Status 'Журнал скопирован в буфер обмена'
            }
        } catch {
            Show-WarnBox ('Не удалось скопировать журнал: ' + $_.Exception.Message)
        }
    })

    $btnSaveLog.Add_Click({
        try {
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = 'Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*'
            $sfd.FileName = 'FS-AccessManager_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log'
            if ($sfd.ShowDialog($script:Form) -eq [System.Windows.Forms.DialogResult]::OK) {
                [System.IO.File]::WriteAllText($sfd.FileName, $script:LogBox.Text, (New-Object System.Text.UTF8Encoding($true)))
                Write-UiLog ('Журнал сохранён: ' + $sfd.FileName) 'Ok'
            }
        } catch {
            Show-WarnBox ('Не удалось сохранить журнал: ' + $_.Exception.Message)
        }
    })

    # Первичная инициализация после появления окна: домен, сохранённый путь, предупреждения.
    $form.Add_Shown({
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
            Write-UiLog 'Поток не STA: копирование в буфер обмена может не работать. Запускайте powershell.exe -STA (или pwsh -STA).' 'Warn'
        }
        Write-UiLog ('Запуск FS-AccessManager. Пользователь: ' + $env:USERDOMAIN + '\' + $env:USERNAME)
        if (-not $script:State.ADReady) {
            Write-UiLog 'Модуль ActiveDirectory не загружен: операции с AD недоступны.' 'Error'
            Show-ErrorBox (Join-Lines @(
                'Не удалось загрузить модуль ActiveDirectory.',
                '',
                'Установите средства RSAT:',
                '  Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0',
                'или',
                '  Install-WindowsFeature RSAT-AD-PowerShell',
                '',
                'Текст ошибки:',
                $script:ModuleError))
        }
        Initialize-DomainInfo
        if ($script:State.RootPath) {
            $script:txtRoot.Text = $script:State.RootPath
            if (Test-Path -LiteralPath $script:State.RootPath -PathType Container) {
                Set-CurrentPath $script:State.RootPath
            } else {
                Write-UiLog ('Сохранённая корневая папка недоступна: ' + $script:State.RootPath) 'Warn'
            }
        }
    })

    # При закрытии окна сохраняем параметры подключения в JSON.
    $form.Add_FormClosing({
        try { Save-Settings } catch { }
    })
}

#endregion


#region ---------- Запуск ----------
# Порядок запуска: импорт модуля ActiveDirectory, построение окна,
# применение сохранённых настроек, запуск цикла сообщений WinForms.

if (-not ('System.Windows.Forms.Form' -as [type])) {
    Add-Type -AssemblyName System.Windows.Forms
}

try {
    # Без модуля ActiveDirectory скрипт запустится, но операции с AD будут недоступны.
    Import-Module ActiveDirectory -ErrorAction Stop
    $script:State.ADReady = $true
} catch {
    $script:State.ADReady = $false
    $script:ModuleError = $_.Exception.Message
}

Build-UI
Initialize-Settings

try {
    [void][System.Windows.Forms.Application]::Run($script:Form)
} catch {
    # последняя линия обороны: показать ошибку, если что-то упало вне обработчиков
    [void][System.Windows.Forms.MessageBox]::Show(
        ('Непредвиденная ошибка:' + [Environment]::NewLine + $_.Exception.Message + [Environment]::NewLine + $_.ScriptStackTrace),
        'FS-AccessManager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
}

#endregion
