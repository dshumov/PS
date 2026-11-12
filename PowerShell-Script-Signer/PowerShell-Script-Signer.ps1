<#
.SYNOPSIS
    GUI-утилита для подписания PowerShell-скриптов сертификатом (Authenticode).

.DESCRIPTION
    Позволяет выбрать действующий сертификат подписи кода из хранилищ
    CurrentUser\My и LocalMachine\My, подписать выбранный .ps1-файл
    алгоритмом SHA-256 и сразу проверить подпись. Подписание и проверка
    выполняются стандартными командлетами Set-AuthenticodeSignature и
    Get-AuthenticodeSignature, поэтому файл остаётся обычным
    подписанным PowerShell-скриптом.

.NOTES
    Требования:
      * Windows PowerShell 5.1 или PowerShell 7+ (Windows);
      * .NET Framework с System.Windows.Forms;
      * сертификат с закрытым ключом и назначением «Подписание кода»
        (Code Signing, OID 1.3.6.1.5.5.7.3.3).

    Чтобы подпись считалась доверенной на других компьютерах, сертификат
    издателя (или корневой сертификат вашего УЦ) должен быть установлен
    в хранилище «Доверенные корневые центры сертификации».
#>

# Требуется Windows PowerShell 5.1+ (WinForms и Authenticode).
#requires -Version 5.1

# ----------------------------------------------------------------------------
#  ИНИЦИАЛИЗАЦИЯ: WinForms, строгий режим и состояние приложения
# ----------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Включаем визуальные стили Windows для элементов управления.
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# StrictMode помогает ловить опечатки в именах свойств и переменных.
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Состояние GUI: выбранный сертификат, его хранилище и путь к скрипту.
$script:SelectedCertificate = $null
$script:SelectedCertificateSource = $null
$script:SelectedScriptPath = $null

# ============================================================================
#  РАБОТА С СЕРТИФИКАТАМИ
# ============================================================================

# Проверяет, что сертификат пригоден для подписи кода:
#   * есть закрытый ключ (без него подписать нельзя);
#   * срок действия уже начался и ещё не истёк;
#   * в Enhanced Key Usage есть OID 1.3.6.1.5.5.7.3.3 (Code Signing)
#     либо расширение EKU отсутствует (сертификат без ограничения назначения).
# Возвращает $true или $false.

function Test-CodeSigningCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    if (-not $Certificate.HasPrivateKey) {
        return $false
    }

    $now = Get-Date
    if (($Certificate.NotBefore -gt $now) -or ($Certificate.NotAfter -lt $now)) {
        return $false
    }

    # OID назначения Enhanced Key Usage "Code Signing".
    $codeSigningOid = '1.3.6.1.5.5.7.3.3'
    $ekuExtension = $null

    foreach ($extension in $Certificate.Extensions) {
        if ($extension -is [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]) {
            $ekuExtension = $extension
            break
        }
    }

    # Отсутствие EKU означает, что сертификат не ограничен конкретным назначением.
    if ($null -eq $ekuExtension) {
        return $true
    }

    foreach ($oid in $ekuExtension.EnhancedKeyUsages) {
        if ($oid.Value -eq $codeSigningOid) {
            return $true
        }
    }

    return $false
}

# Перебирает хранилища сертификатов CurrentUser\My и LocalMachine\My,
# отбирает пригодные для подписи кода и возвращает список объектов
# с полями для таблицы: владелец, издатель, хранилище, срок, отпечаток.
# Копии сертификатов создаются через new(X509Certificate2), чтобы они
# оставались доступны после закрытия хранилища.
function Get-CodeSigningCertificates {
    $result = [System.Collections.Generic.List[System.Object]]::new()
    $locations = @(
        [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser,
        [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
    )

    foreach ($location in $locations) {
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new(
            'My',
            $location
        )

        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

            foreach ($certificate in $store.Certificates) {
                if (-not (Test-CodeSigningCertificate -Certificate $certificate)) {
                    continue
                }

                $subjectName = $certificate.GetNameInfo(
                    [System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName,
                    $false
                )
                if ([string]::IsNullOrWhiteSpace($subjectName)) {
                    $subjectName = $certificate.Subject
                }

                $issuerName = $certificate.GetNameInfo(
                    [System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName,
                    $true
                )
                if ([string]::IsNullOrWhiteSpace($issuerName)) {
                    $issuerName = $certificate.Issuer
                }

                # Создаем отдельный объект сертификата, чтобы он оставался пригоден
                # после закрытия хранилища.
                $certificateCopy = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                    $certificate
                )

                $result.Add([pscustomobject]@{
                    Certificate = $certificateCopy
                    Name        = $subjectName
                    Issuer      = $issuerName
                    Store       = ('{0}\My' -f $location)
                    NotBefore   = $certificate.NotBefore
                    NotAfter    = $certificate.NotAfter
                    Thumbprint  = $certificate.Thumbprint
                })
            }
        }
        catch {
            # Например, чтение LocalMachine может быть ограничено политикой.
            # Остальные доступные хранилища все равно будут обработаны.
        }
        finally {
            $store.Close()
            $store.Dispose()
        }
    }

    return @($result | Sort-Object Name, NotAfter)
}

# Показывает модальное окно со списком сертификатов.
# Возвращает выбранный элемент (или $null, если пользователь нажал «Отмена»).
# Невыбранные сертификаты освобождаются, чтобы не удерживать закрытые ключи.
function Show-CertificatePicker {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.IWin32Window]$Owner
    )

    $certificates = @(Get-CodeSigningCertificates)

    if ($certificates.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $Owner,
            "Не найден действующий сертификат с закрытым ключом и назначением `"Подписание кода`".`r`n`r`nПроверены хранилища CurrentUser\My и LocalMachine\My.",
            'Сертификаты не найдены',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $null
    }

    $picker = New-Object System.Windows.Forms.Form
    $picker.Text = 'Выбор сертификата для подписания кода'
    $picker.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $picker.Size = [System.Drawing.Size]::new(1040, 500)
    $picker.MinimumSize = [System.Drawing.Size]::new(850, 400)
    $picker.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $picker.MaximizeBox = $true
    $picker.MinimizeBox = $false
    $picker.ShowInTaskbar = $false
    $picker.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

    $label = New-Object System.Windows.Forms.Label
    $label.Location = [System.Drawing.Point]::new(12, 12)
    $label.Size = [System.Drawing.Size]::new(990, 34)
    $label.Text = 'Показаны действующие сертификаты с закрытым ключом и назначением подписания кода. Выберите строку и нажмите «Выбрать».'
    $label.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                    [System.Windows.Forms.AnchorStyles]::Left -bor `
                    [System.Windows.Forms.AnchorStyles]::Right
    $picker.Controls.Add($label)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = [System.Drawing.Point]::new(12, 50)
    $grid.Size = [System.Drawing.Size]::new(1000, 350)
    $grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                  [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                  [System.Windows.Forms.AnchorStyles]::Left -bor `
                  [System.Windows.Forms.AnchorStyles]::Right
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.ReadOnly = $true
    $grid.MultiSelect = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.BackgroundColor = [System.Drawing.Color]::White

    [void]$grid.Columns.Add('Name', 'Владелец')
    [void]$grid.Columns.Add('Issuer', 'Издатель')
    [void]$grid.Columns.Add('Store', 'Хранилище')
    [void]$grid.Columns.Add('NotAfter', 'Действителен до')
    [void]$grid.Columns.Add('Thumbprint', 'Отпечаток')

    $grid.Columns['Name'].FillWeight = 135
    $grid.Columns['Issuer'].FillWeight = 115
    $grid.Columns['Store'].FillWeight = 65
    $grid.Columns['NotAfter'].FillWeight = 65
    $grid.Columns['Thumbprint'].FillWeight = 145

    foreach ($item in $certificates) {
        $rowIndex = $grid.Rows.Add(
            $item.Name,
            $item.Issuer,
            $item.Store,
            $item.NotAfter.ToString('dd.MM.yyyy HH:mm'),
            $item.Thumbprint
        )
        $grid.Rows[$rowIndex].Tag = $item
    }

    if ($grid.Rows.Count -gt 0) {
        $grid.Rows[0].Selected = $true
    }

    $picker.Controls.Add($grid)

    $buttonSelect = New-Object System.Windows.Forms.Button
    $buttonSelect.Text = 'Выбрать'
    $buttonSelect.Size = [System.Drawing.Size]::new(120, 32)
    $buttonSelect.Location = [System.Drawing.Point]::new(758, 414)
    $buttonSelect.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                           [System.Windows.Forms.AnchorStyles]::Right
    $picker.Controls.Add($buttonSelect)

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = 'Отмена'
    $buttonCancel.Size = [System.Drawing.Size]::new(120, 32)
    $buttonCancel.Location = [System.Drawing.Point]::new(892, 414)
    $buttonCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                           [System.Windows.Forms.AnchorStyles]::Right
    $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $picker.Controls.Add($buttonCancel)

    $selectCertificate = {
        if ($grid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                $picker,
                'Сначала выберите сертификат в таблице.',
                'Сертификат не выбран',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $picker.Tag = $grid.SelectedRows[0].Tag
        $picker.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $picker.Close()
    }.GetNewClosure()

    $buttonSelect.Add_Click($selectCertificate)

    $doubleClickHandler = {
        param($sender, $eventArgs)

        if ($eventArgs.RowIndex -ge 0) {
            $grid.ClearSelection()
            $grid.Rows[$eventArgs.RowIndex].Selected = $true
            & $selectCertificate
        }
    }.GetNewClosure()
    $grid.Add_CellDoubleClick($doubleClickHandler)

    $picker.AcceptButton = $buttonSelect
    $picker.CancelButton = $buttonCancel

    $dialogResult = $picker.ShowDialog($Owner)
    $selection = $null

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $selection = $picker.Tag
    }

    # Освобождаем объекты невыбранных сертификатов.
    foreach ($item in $certificates) {
        if (($null -eq $selection) -or
            -not [object]::ReferenceEquals($item.Certificate, $selection.Certificate)) {
            $item.Certificate.Dispose()
        }
    }

    $picker.Dispose()
    return $selection
}

# ============================================================================
#  ГЛАВНОЕ ОКНО
#  Блоки интерфейса:
#    1) выбор сертификата;
#    2) выбор PowerShell-скрипта;
#    3–4) подписание и проверка подписи;
#    журнал операций и кнопка «Закрыть».
# ============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Подписание PowerShell-скриптов'
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Size = [System.Drawing.Size]::new(980, 700)
$form.MinimumSize = [System.Drawing.Size]::new(900, 640)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

$title = New-Object System.Windows.Forms.Label
$title.Location = [System.Drawing.Point]::new(20, 15)
$title.Size = [System.Drawing.Size]::new(920, 34)
$title.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                [System.Windows.Forms.AnchorStyles]::Left -bor `
                [System.Windows.Forms.AnchorStyles]::Right
$title.Font = [System.Drawing.Font]::new('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$title.Text = 'Подписание PowerShell-скриптов сертификатом'
$form.Controls.Add($title)

$certificateGroup = New-Object System.Windows.Forms.GroupBox
$certificateGroup.Text = '1. Сертификат'
$certificateGroup.Location = [System.Drawing.Point]::new(20, 58)
$certificateGroup.Size = [System.Drawing.Size]::new(925, 145)
$certificateGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                           [System.Windows.Forms.AnchorStyles]::Left -bor `
                           [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($certificateGroup)

$certificateText = New-Object System.Windows.Forms.TextBox
$certificateText.Location = [System.Drawing.Point]::new(15, 25)
$certificateText.Size = [System.Drawing.Size]::new(720, 102)
$certificateText.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                          [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                          [System.Windows.Forms.AnchorStyles]::Left -bor `
                          [System.Windows.Forms.AnchorStyles]::Right
$certificateText.Multiline = $true
$certificateText.ReadOnly = $true
$certificateText.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$certificateText.BackColor = [System.Drawing.Color]::White
$certificateText.Text = 'Сертификат не выбран.'
$certificateGroup.Controls.Add($certificateText)

$certificateButton = New-Object System.Windows.Forms.Button
$certificateButton.Location = [System.Drawing.Point]::new(752, 25)
$certificateButton.Size = [System.Drawing.Size]::new(155, 38)
$certificateButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                            [System.Windows.Forms.AnchorStyles]::Right
$certificateButton.Text = 'Выбрать...'
$certificateGroup.Controls.Add($certificateButton)

$certificateHint = New-Object System.Windows.Forms.Label
$certificateHint.Location = [System.Drawing.Point]::new(752, 72)
$certificateHint.Size = [System.Drawing.Size]::new(155, 55)
$certificateHint.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                          [System.Windows.Forms.AnchorStyles]::Right
$certificateHint.Text = "CurrentUser\My`r`nLocalMachine\My"
$certificateHint.ForeColor = [System.Drawing.Color]::DimGray
$certificateGroup.Controls.Add($certificateHint)

$scriptGroup = New-Object System.Windows.Forms.GroupBox
$scriptGroup.Text = '2. PowerShell-скрипт'
$scriptGroup.Location = [System.Drawing.Point]::new(20, 215)
$scriptGroup.Size = [System.Drawing.Size]::new(925, 94)
$scriptGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                      [System.Windows.Forms.AnchorStyles]::Left -bor `
                      [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($scriptGroup)

$scriptText = New-Object System.Windows.Forms.TextBox
$scriptText.Location = [System.Drawing.Point]::new(15, 35)
$scriptText.Size = [System.Drawing.Size]::new(720, 25)
$scriptText.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                     [System.Windows.Forms.AnchorStyles]::Left -bor `
                     [System.Windows.Forms.AnchorStyles]::Right
$scriptText.ReadOnly = $true
$scriptText.BackColor = [System.Drawing.Color]::White
$scriptGroup.Controls.Add($scriptText)

$scriptButton = New-Object System.Windows.Forms.Button
$scriptButton.Location = [System.Drawing.Point]::new(752, 29)
$scriptButton.Size = [System.Drawing.Size]::new(155, 38)
$scriptButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                       [System.Windows.Forms.AnchorStyles]::Right
$scriptButton.Text = 'Выбрать...'
$scriptGroup.Controls.Add($scriptButton)

$actionGroup = New-Object System.Windows.Forms.GroupBox
$actionGroup.Text = '3–4. Подписание и проверка'
$actionGroup.Location = [System.Drawing.Point]::new(20, 321)
$actionGroup.Size = [System.Drawing.Size]::new(925, 105)
$actionGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                      [System.Windows.Forms.AnchorStyles]::Left -bor `
                      [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($actionGroup)

$signButton = New-Object System.Windows.Forms.Button
$signButton.Location = [System.Drawing.Point]::new(15, 30)
$signButton.Size = [System.Drawing.Size]::new(190, 45)
$signButton.Text = 'Подписать скрипт'
$signButton.Enabled = $false
$actionGroup.Controls.Add($signButton)

$verifyButton = New-Object System.Windows.Forms.Button
$verifyButton.Location = [System.Drawing.Point]::new(220, 30)
$verifyButton.Size = [System.Drawing.Size]::new(190, 45)
$verifyButton.Text = 'Проверить подпись'
$verifyButton.Enabled = $false
$actionGroup.Controls.Add($verifyButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = [System.Drawing.Point]::new(430, 30)
$statusLabel.Size = [System.Drawing.Size]::new(475, 45)
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                      [System.Windows.Forms.AnchorStyles]::Left -bor `
                      [System.Windows.Forms.AnchorStyles]::Right
$statusLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusLabel.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$statusLabel.Text = 'Выберите сертификат и скрипт.'
$statusLabel.BackColor = [System.Drawing.Color]::WhiteSmoke
$actionGroup.Controls.Add($statusLabel)

$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = 'Журнал'
$logGroup.Location = [System.Drawing.Point]::new(20, 438)
$logGroup.Size = [System.Drawing.Size]::new(925, 185)
$logGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                   [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                   [System.Windows.Forms.AnchorStyles]::Left -bor `
                   [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($logGroup)

$logText = New-Object System.Windows.Forms.RichTextBox
$logText.Location = [System.Drawing.Point]::new(15, 25)
$logText.Size = [System.Drawing.Size]::new(892, 143)
$logText.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                  [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                  [System.Windows.Forms.AnchorStyles]::Left -bor `
                  [System.Windows.Forms.AnchorStyles]::Right
$logText.ReadOnly = $true
$logText.BackColor = [System.Drawing.Color]::White
$logText.Font = [System.Drawing.Font]::new('Consolas', 9)
$logGroup.Controls.Add($logText)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = [System.Drawing.Point]::new(810, 632)
$closeButton.Size = [System.Drawing.Size]::new(135, 32)
$closeButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                      [System.Windows.Forms.AnchorStyles]::Right
$closeButton.Text = 'Закрыть'
$form.Controls.Add($closeButton)

# Добавляет строку в журнал окна. Цвет зависит от уровня:
# INFO — чёрный, OK — зелёный, WARN — оранжевый, ERROR — красный.
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'OK'    { [System.Drawing.Color]::DarkGreen }
        'WARN'  { [System.Drawing.Color]::DarkOrange }
        'ERROR' { [System.Drawing.Color]::DarkRed }
        default { [System.Drawing.Color]::Black }
    }

    $line = '{0:HH:mm:ss} [{1}] {2}{3}' -f (Get-Date), $Level, $Message, [Environment]::NewLine
    $logText.SelectionStart = $logText.TextLength
    $logText.SelectionLength = 0
    $logText.SelectionColor = $color
    $logText.AppendText($line)
    $logText.SelectionColor = $logText.ForeColor
    $logText.ScrollToCaret()
}

# Обновляет цветную строку состояния под кнопками (нейтрально/инфо/успех/
# предупреждение/ошибка).
function Set-Status {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [ValidateSet('Neutral', 'Info', 'Success', 'Warning', 'Error')]
        [string]$Kind = 'Neutral'
    )

    $statusLabel.Text = $Text

    switch ($Kind) {
        'Info' {
            $statusLabel.BackColor = [System.Drawing.Color]::LightCyan
            $statusLabel.ForeColor = [System.Drawing.Color]::MidnightBlue
        }
        'Success' {
            $statusLabel.BackColor = [System.Drawing.Color]::Honeydew
            $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        }
        'Warning' {
            $statusLabel.BackColor = [System.Drawing.Color]::LemonChiffon
            $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        }
        'Error' {
            $statusLabel.BackColor = [System.Drawing.Color]::MistyRose
            $statusLabel.ForeColor = [System.Drawing.Color]::DarkRed
        }
        default {
            $statusLabel.BackColor = [System.Drawing.Color]::WhiteSmoke
            $statusLabel.ForeColor = [System.Drawing.Color]::Black
        }
    }
}

# Включает/выключает кнопки действий:
#   «Проверить подпись» — если выбран существующий файл;
#   «Подписать скрипт»  — если выбран файл и сертификат.
function Update-ActionButtons {
    $scriptExists = (-not [string]::IsNullOrWhiteSpace($script:SelectedScriptPath)) -and
                    (Test-Path -LiteralPath $script:SelectedScriptPath -PathType Leaf)

    $verifyButton.Enabled = $scriptExists
    $signButton.Enabled = $scriptExists -and ($null -ne $script:SelectedCertificate)
}

# Разбирает результат Get-AuthenticodeSignature и выводит его в строку
# состояния и журнал. Обрабатываются статусы:
#   Valid             — подпись действительна;
#   NotSigned         — подписи нет;
#   HashMismatch      — файл изменён после подписания;
#   NotTrusted        — подпись есть, но цепочка сертификата не доверена;
#   NotSupportedFileFormat — формат файла не поддерживает Authenticode;
#   Incompatible      — подпись несовместима с системой/PowerShell;
#   прочие            — выводятся как есть с предупреждением.
# -AfterSigning добавляет к сообщению приставку «Результат после подписания».
function Show-SignatureResult {
    param(
        [Parameter(Mandatory = $true)]
        $Signature,

        [switch]$AfterSigning
    )

    $statusName = $Signature.Status.ToString()
    $prefix = if ($AfterSigning) { 'Результат после подписания: ' } else { '' }

    switch ($statusName) {
        'Valid' {
            $message = $prefix + 'скрипт подписан, подпись действительна.'
            Set-Status -Text 'ПОДПИСАН: подпись действительна' -Kind Success
            Write-Log -Message $message -Level OK
        }
        'NotSigned' {
            $message = $prefix + 'скрипт не подписан.'
            Set-Status -Text 'НЕ ПОДПИСАН' -Kind Warning
            Write-Log -Message $message -Level WARN
        }
        'HashMismatch' {
            $message = $prefix + 'подпись присутствует, но содержимое скрипта было изменено после подписания.'
            Set-Status -Text 'ПОДПИСЬ НЕДЕЙСТВИТЕЛЬНА: файл изменен' -Kind Error
            Write-Log -Message $message -Level ERROR
        }
        'NotTrusted' {
            $message = $prefix + 'скрипт подписан, но цепочка сертификата не является доверенной на этом компьютере.'
            Set-Status -Text 'ПОДПИСАН, но сертификат не доверен' -Kind Warning
            Write-Log -Message $message -Level WARN
        }
        'NotSupportedFileFormat' {
            $message = $prefix + 'формат файла не поддерживает Authenticode-подпись.'
            Set-Status -Text 'Формат файла не поддерживается' -Kind Error
            Write-Log -Message $message -Level ERROR
        }
        'Incompatible' {
            $message = $prefix + 'подпись несовместима с текущей системой или версией PowerShell.'
            Set-Status -Text 'Подпись несовместима' -Kind Error
            Write-Log -Message $message -Level ERROR
        }
        default {
            if ($null -ne $Signature.SignerCertificate) {
                $message = $prefix + ('подпись присутствует, но PowerShell вернул статус {0}.' -f $statusName)
                Set-Status -Text ('ПОДПИСАН, статус: {0}' -f $statusName) -Kind Warning
                Write-Log -Message $message -Level WARN
            }
            else {
                $message = $prefix + ('не удалось подтвердить наличие корректной подписи. Статус: {0}.' -f $statusName)
                Set-Status -Text ('Ошибка проверки: {0}' -f $statusName) -Kind Error
                Write-Log -Message $message -Level ERROR
            }
        }
    }

    if ($null -ne $Signature.SignerCertificate) {
        $signerName = $Signature.SignerCertificate.GetNameInfo(
            [System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName,
            $false
        )
        if ([string]::IsNullOrWhiteSpace($signerName)) {
            $signerName = $Signature.SignerCertificate.Subject
        }

        Write-Log -Message ('Сертификат подписи: {0}; отпечаток: {1}' -f `
            $signerName,
            $Signature.SignerCertificate.Thumbprint) -Level INFO
    }

    if (-not [string]::IsNullOrWhiteSpace($Signature.StatusMessage)) {
        Write-Log -Message ('Сообщение PowerShell: {0}' -f $Signature.StatusMessage.Trim()) -Level INFO
    }
}

# Проверяет подпись уже выбранного файла и возвращает объект подписи
# (его же использует обработчик кнопки «Проверить подпись»).
function Verify-SelectedScript {
    if ([string]::IsNullOrWhiteSpace($script:SelectedScriptPath) -or
        -not (Test-Path -LiteralPath $script:SelectedScriptPath -PathType Leaf)) {
        throw 'Сначала выберите существующий PowerShell-скрипт.'
    }

    Write-Log -Message ('Проверка подписи: {0}' -f $script:SelectedScriptPath) -Level INFO
    $signature = Get-AuthenticodeSignature -LiteralPath $script:SelectedScriptPath -ErrorAction Stop
    Show-SignatureResult -Signature $signature
    return $signature
}

# ============================================================================
#  ОБРАБОТЧИКИ СОБЫТИЙ
# ============================================================================

# Выбор сертификата: открывает окно выбора и запоминает сертификат
# и его хранилище в состоянии скрипта.
$certificateButton.Add_Click({
    try {
        $selected = Show-CertificatePicker -Owner $form
        if ($null -eq $selected) {
            return
        }

        if ($null -ne $script:SelectedCertificate) {
            $script:SelectedCertificate.Dispose()
        }

        $script:SelectedCertificate = $selected.Certificate
        $script:SelectedCertificateSource = $selected.Store

        $certificateText.Text = @(
            'Владелец: {0}' -f $selected.Name
            'Издатель: {0}' -f $selected.Issuer
            'Хранилище: {0}' -f $selected.Store
            'Срок действия: {0:dd.MM.yyyy HH:mm} — {1:dd.MM.yyyy HH:mm}' -f $selected.NotBefore, $selected.NotAfter
            'Отпечаток: {0}' -f $selected.Thumbprint
        ) -join [Environment]::NewLine

        Write-Log -Message ('Выбран сертификат: {0} ({1})' -f $selected.Name, $selected.Thumbprint) -Level OK
        Set-Status -Text 'Сертификат выбран' -Kind Info
        Update-ActionButtons
    }
    catch {
        Set-Status -Text 'Ошибка выбора сертификата' -Kind Error
        Write-Log -Message $_.Exception.Message -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $_.Exception.Message,
            'Ошибка',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

# Выбор скрипта: диалог открытия файла, затем сразу проверяем
# исходное состояние подписи выбранного файла.
$scriptButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Выберите PowerShell-скрипт'
    $dialog.Filter = 'PowerShell-скрипты (*.ps1)|*.ps1|Все файлы (*.*)|*.*'
    $dialog.FilterIndex = 1
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false
    $dialog.RestoreDirectory = $true

    try {
        if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $script:SelectedScriptPath = $dialog.FileName
        $scriptText.Text = $script:SelectedScriptPath
        Write-Log -Message ('Выбран скрипт: {0}' -f $script:SelectedScriptPath) -Level OK
        Set-Status -Text 'Скрипт выбран' -Kind Info
        Update-ActionButtons

        # Сразу показываем исходный статус подписи выбранного файла.
        try {
            Verify-SelectedScript | Out-Null
        }
        catch {
            Write-Log -Message ('Не удалось выполнить первичную проверку: {0}' -f $_.Exception.Message) -Level WARN
        }
    }
    finally {
        $dialog.Dispose()
    }
})

# Подписание: Set-AuthenticodeSignature добавляет в конец файла блок
# подписи, после чего файл проверяется заново уже с диска.
$signButton.Add_Click({
    $signButton.Enabled = $false
    $verifyButton.Enabled = $false
    $form.UseWaitCursor = $true

    try {
        if ($null -eq $script:SelectedCertificate) {
            throw 'Сначала выберите сертификат.'
        }

        if ([string]::IsNullOrWhiteSpace($script:SelectedScriptPath) -or
            -not (Test-Path -LiteralPath $script:SelectedScriptPath -PathType Leaf)) {
            throw 'Сначала выберите существующий PowerShell-скрипт.'
        }

        if (-not $script:SelectedCertificate.HasPrivateKey) {
            throw 'У выбранного сертификата нет доступного закрытого ключа.'
        }

        Write-Log -Message ('Подписание файла алгоритмом SHA-256: {0}' -f $script:SelectedScriptPath) -Level INFO

        $signResult = Set-AuthenticodeSignature `
            -LiteralPath $script:SelectedScriptPath `
            -Certificate $script:SelectedCertificate `
            -HashAlgorithm SHA256 `
            -ErrorAction Stop

        if ($null -eq $signResult.SignerCertificate) {
            throw ('PowerShell не добавил подпись. Статус: {0}. {1}' -f `
                $signResult.Status,
                $signResult.StatusMessage)
        }

        # Выполняем отдельную проверку уже записанного файла.
        $verification = Get-AuthenticodeSignature `
            -LiteralPath $script:SelectedScriptPath `
            -ErrorAction Stop

        Show-SignatureResult -Signature $verification -AfterSigning

        if ($verification.Status.ToString() -eq 'Valid') {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                'Скрипт успешно подписан. Подпись действительна.',
                'Подписание завершено',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        elseif ($null -ne $verification.SignerCertificate) {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                ("Подпись добавлена, но результат проверки: {0}.`r`n`r`n{1}" -f `
                    $verification.Status,
                    $verification.StatusMessage),
                'Подпись добавлена с предупреждением',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
        else {
            throw ('После подписания подпись не обнаружена. Статус: {0}.' -f $verification.Status)
        }
    }
    catch {
        Set-Status -Text 'Ошибка подписания' -Kind Error
        Write-Log -Message $_.Exception.Message -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $_.Exception.Message,
            'Ошибка подписания',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $form.UseWaitCursor = $false
        Update-ActionButtons
    }
})

# Проверка подписи выбранного файла по кнопке.
$verifyButton.Add_Click({
    $signButton.Enabled = $false
    $verifyButton.Enabled = $false
    $form.UseWaitCursor = $true

    try {
        $signature = Verify-SelectedScript

        $icon = if ($signature.Status.ToString() -eq 'Valid') {
            [System.Windows.Forms.MessageBoxIcon]::Information
        }
        elseif ($signature.Status.ToString() -eq 'NotSigned') {
            [System.Windows.Forms.MessageBoxIcon]::Warning
        }
        else {
            [System.Windows.Forms.MessageBoxIcon]::Warning
        }

        $message = switch ($signature.Status.ToString()) {
            'Valid'        { 'Скрипт подписан. Подпись действительна.' }
            'NotSigned'    { 'Скрипт не подписан.' }
            'HashMismatch' { 'Подпись присутствует, но скрипт был изменен после подписания.' }
            'NotTrusted'   { 'Скрипт подписан, но сертификат не является доверенным на этом компьютере.' }
            default        { "Результат проверки: {0}.`r`n`r`n{1}" -f $signature.Status, $signature.StatusMessage }
        }

        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $message,
            'Проверка подписи',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $icon
        ) | Out-Null
    }
    catch {
        Set-Status -Text 'Ошибка проверки подписи' -Kind Error
        Write-Log -Message $_.Exception.Message -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $_.Exception.Message,
            'Ошибка проверки',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $form.UseWaitCursor = $false
        Update-ActionButtons
    }
})

# Закрытие окна и освобождение закрытого ключа выбранного сертификата.
$closeButton.Add_Click({
    $form.Close()
})

$form.Add_FormClosed({
    if ($null -ne $script:SelectedCertificate) {
        $script:SelectedCertificate.Dispose()
        $script:SelectedCertificate = $null
    }
})

# Первые сообщения в журнале и запуск модального цикла сообщений.
# ShowDialog блокирует консоль до закрытия окна.
Write-Log -Message 'Приложение запущено.' -Level INFO
Write-Log -Message 'Выберите сертификат и PowerShell-скрипт.' -Level INFO

[void]$form.ShowDialog()
$form.Dispose()
