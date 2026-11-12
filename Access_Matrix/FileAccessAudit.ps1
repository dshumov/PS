<#
.SYNOPSIS
    Аудит и матрица доступа к файловому ресурсу (NTFS ACL) с расшифровкой через Active Directory.

.DESCRIPTION
    Скрипт обходит дерево каталогов (по умолчанию — 3 уровня вложенности), читает
    NTFS-права (ACL) каждой папки, сопоставляет найденные учётные записи с
    пользователями и группами Active Directory и рекурсивно разворачивает членство
    в группах. На выходе формируются:

      * AccessMatrix_<дата>.csv    — матрица «пользователь x папка x права»;
      * GroupMembership_<дата>.csv — членство пользователей в группах;
      * AccessMatrix_<дата>.html   — интерактивный отчёт с фильтрами и экспортом;
      * SearchUserAccess.ps1       — вспомогательный скрипт поиска по пользователю.

    Скрипт только читает права: ни ACL, ни данные на диске он не изменяет.

.EXAMPLE
    .\FileAccessAudit.ps1

.NOTES
    Требования:
      * Windows PowerShell 5.1 или PowerShell 7+;
      * модуль ActiveDirectory (RSAT) и доступ к контроллеру домена;
      * права на чтение ACL сканируемых папок.

    Перед запуском заполните блок «НАСТРОЙКИ»: путь к ресурсу, каталог отчётов и
    distinguishedName (DN) подразделений (OU) Active Directory.
#>

# ============================================================================
#  НАСТРОЙКИ — единственный блок, который нужно изменить перед запуском
# ============================================================================

# Каталог самого скрипта. В него сохраняется сгенерированный SearchUserAccess.ps1.
# Обычно менять не нужно: $PSScriptRoot — папка, из которой запущен скрипт.
$scriptDir = $PSScriptRoot

# Каталог для готовых отчётов (CSV/HTML). Создаётся автоматически.
# ЧТО ВСТАВИТЬ: любой путь, доступный на запись.
# Пример: 'D:\Reports\AccessMatrix\log' или '\\fileserver\reports\AccessMatrix'
$outputDir = Join-Path $scriptDir 'log'

# Корень файлового ресурса, который нужно просканировать.
# ЧТО ВСТАВИТЬ: локальный или UNC-путь к общей папке.
# Пример: 'D:\FileShare' или '\\fileserver\share'
$scanPath = 'D:\FileShare'

# OU (подразделение AD), в которой лежат учётные записи пользователей.
# ЧТО ВСТАВИТЬ: distinguishedName вашего OU.
# Пример: 'OU=Users,OU=Company,DC=example,DC=local'
# Где взять DN: оснастка «Пользователи и компьютеры Active Directory» →
# включить «Дополнительные компоненты» → свойства объекта → вкладка
# «Редактор атрибутов» → атрибут distinguishedName.
$usersOU = 'OU=Users,OU=Company,DC=example,DC=local'

# OU, в которой лежат группы доступа к файловому серверу.
# ЧТО ВСТАВИТЬ: distinguishedName вашего OU с группами.
# Пример: 'OU=FileServers,OU=Resources,OU=Company,DC=example,DC=local'
$groupsOU = 'OU=FileServers,OU=Resources,OU=Company,DC=example,DC=local'

# Пути, которые нужно исключить из сканирования (например, дистрибутивы,
# драйверы, архив). Сравнение идёт по началу пути (-like "<путь>*"), то есть
# исключается папка вместе со всем содержимым.
# ЧТО ВСТАВИТЬ: свои пути. Если исключения не нужны — оставьте @().
$excludePaths = @(
    'D:\FileShare\Infrastructure\Distributives'   # ПРИМЕР — замените на свой путь
    'D:\FileShare\Infrastructure\Drivers'         # ПРИМЕР — замените на свой путь
)

# ----------------------------------------------------------------------------
# ЭТАП 0. Подготовка: создаём каталог отчётов, если его ещё нет
# ----------------------------------------------------------------------------
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force
}

# Словарь перевода имён прав NTFS (FileSystemRights) на русский язык.
# Если в ACL встретится право без перевода — оно выводится как есть.
$rightsTranslation = @{
    "FullControl" = "Полный доступ"
    "Modify" = "Изменение" 
    "ReadAndExecute" = "Чтение и выполнение"
    "ListDirectory" = "Чтение каталога"
    "Read" = "Чтение"
    "Write" = "Запись"
    "Delete" = "Удаление"
    "ChangePermissions" = "Изменение прав"
    "TakeOwnership" = "Владение"
}

# Проверяет, попадает ли путь под исключение из $excludePaths.
# Сравнение — по шаблону "<исключённый путь>*": исключается папка со всем содержимым.
# Возвращает $true, если путь нужно пропустить.
function Test-ExcludedPath {
    param([string]$Path)
    
    foreach ($excludedPath in $excludePaths) {
        if ($Path -like "$excludedPath*") {
            return $true
        }
    }
    return $false
}

# Рекурсивно собирает список папок, начиная с $Path.
#   MaxDepth     — максимальная глубина вложенности (0 = только корень);
#   CurrentDepth — служебный счётчик рекурсии.
# Исключённые пути и папки, к которым нет доступа, пропускаются без падения скрипта.
function Get-FoldersRecursive {
    param(
        [string]$Path,
        [int]$MaxDepth,
        [int]$CurrentDepth = 0
    )
    
    $folders = @()
    
    # Проверяем, не исключен ли текущий путь
    if (Test-ExcludedPath -Path $Path) {
        Write-Host "Пропуск исключенного пути: $Path" -ForegroundColor Yellow
        return $folders
    }
    
    if ($CurrentDepth -le $MaxDepth) {
        try {
            $currentFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue
            foreach ($folder in $currentFolders) {
                # Проверяем, не исключена ли текущая папка
                if (Test-ExcludedPath -Path $folder.FullName) {
                    Write-Host "Пропуск исключенной папки: $($folder.FullName)" -ForegroundColor Yellow
                    continue
                }
                
                $folders += $folder
                if ($CurrentDepth -lt $MaxDepth) {
                    $subFolders = Get-FoldersRecursive -Path $folder.FullName -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1)
                    $folders += $subFolders
                }
            }
        }
        catch {
            Write-Warning "Ошибка доступа к папке: $Path"
        }
    }
    
    return $folders
}

# Переводит строку прав вида "Read, Write" в "Чтение, Запись",
# используя словарь $rightsTranslation.
function Convert-RightsToRussian {
    param([string]$rights)
    
    $translatedRights = @()
    $rightsArray = $rights -split ", "
    
    foreach ($right in $rightsArray) {
        if ($rightsTranslation.ContainsKey($right.Trim())) {
            $translatedRights += $rightsTranslation[$right.Trim()]
        } else {
            $translatedRights += $right
        }
    }
    
    return ($translatedRights -join ", ")
}

# Читает пользователей из OU $usersOU и строит хэш-таблицу
# "sAMAccountName → { DisplayName, SamAccountName, MemberOf }".
# MemberOf нужен, чтобы затем определить группы пользователя.
function Get-ADUsersInfo {
    try {
        Write-Host "Получение информации о пользователях из Active Directory..." -ForegroundColor Green
        
        # Получаем всех пользователей
        $users = Get-ADUser -Filter * -SearchBase $usersOU -Properties DisplayName, SamAccountName, Name, MemberOf | 
                 Select-Object SamAccountName, DisplayName, Name, MemberOf
        
        # Создаем хэш-таблицу для быстрого поиска
        $usersMap = @{}
        
        foreach ($user in $users) {
            $displayName = if ($user.DisplayName) { $user.DisplayName } else { $user.Name }
            $usersMap[$user.SamAccountName] = @{
                DisplayName = $displayName
                SamAccountName = $user.SamAccountName
                MemberOf = $user.MemberOf
            }
        }
        
        return $usersMap
    }
    catch {
        Write-Warning "Ошибка при получении данных пользователей из AD: $_"
        return @{}
    }
}

# Читает группы из OU $groupsOU и строит хэш-таблицу
# "sAMAccountName → { DisplayName, SamAccountName, MemberOf }".
# MemberOf позволяет отследить вложенность групп (группа внутри группы).
function Get-ADGroupsInfo {
    try {
        Write-Host "Получение информации о группах из Active Directory..." -ForegroundColor Green
        
        # Получаем все группы
        $groups = Get-ADGroup -Filter * -SearchBase $groupsOU -Properties DisplayName, SamAccountName, Name, MemberOf | 
                  Select-Object SamAccountName, DisplayName, Name, MemberOf
        
        # Создаем хэш-таблицу для быстрого поиска
        $groupsMap = @{}
        
        foreach ($group in $groups) {
            $displayName = if ($group.DisplayName) { $group.DisplayName } else { $group.Name }
            $groupsMap[$group.SamAccountName] = @{
                DisplayName = $displayName
                SamAccountName = $group.SamAccountName
                MemberOf = $group.MemberOf
            }
        }
        
        return $groupsMap
    }
    catch {
        Write-Warning "Ошибка при получении данных групп из AD: $_"
        return @{}
    }
}

# Возвращает все группы пользователя, включая вложенные (рекурсивно).
# Из DN вида "CN=Group,OU=..." извлекается имя группы (CN),
# после чего вызывается Get-GroupAllGroups для родительских групп.
# Результат дедублицируется через Select-Object -Unique.
function Get-UserAllGroups {
    param(
        [string]$userSamAccountName,
        [hashtable]$usersInfo,
        [hashtable]$groupsInfo
    )
    
    $allGroups = @()
    $user = $usersInfo[$userSamAccountName]
    
    if ($user -and $user.MemberOf) {
        foreach ($groupDN in $user.MemberOf) {
            # Извлекаем CN из DistinguishedName
            if ($groupDN -match "CN=([^,]+),") {
                $groupName = $matches[1]
                $allGroups += $groupName
                
                # Рекурсивно получаем родительские группы
                $parentGroups = Get-GroupAllGroups -groupSamAccountName $groupName -groupsInfo $groupsInfo
                $allGroups += $parentGroups
            }
        }
    }
    
    return $allGroups | Select-Object -Unique
}

# Рекурсивно поднимается по цепочке MemberOf и возвращает все родительские группы.
# Так учитываются права, выданные не напрямую группе пользователя, а группе,
# в которую эта группа вложена (например, Project-Readers → All-Readers).
function Get-GroupAllGroups {
    param(
        [string]$groupSamAccountName,
        [hashtable]$groupsInfo
    )
    
    $allParentGroups = @()
    $group = $groupsInfo[$groupSamAccountName]
    
    if ($group -and $group.MemberOf) {
        foreach ($parentGroupDN in $group.MemberOf) {
            if ($parentGroupDN -match "CN=([^,]+),") {
                $parentGroupName = $matches[1]
                $allParentGroups += $parentGroupName
                
                # Рекурсивно получаем следующие уровни групп
                $grandParentGroups = Get-GroupAllGroups -groupSamAccountName $parentGroupName -groupsInfo $groupsInfo
                $allParentGroups += $grandParentGroups
            }
        }
    }
    
    return $allParentGroups | Select-Object -Unique
}

# Приводит IdentityReference из ACL к удобному виду:
#   * отбрасывает встроенные учётные записи (SYSTEM, Administrators,
#     CREATOR OWNER и т. п.) — они не интересны в матрице доступа;
#   * убирает доменную приставку: "DOMAIN\user" → "user".
# Возвращает $null, если запись нужно пропустить.
function Format-AccountName {
    param([string]$identity)
    
    # Убираем встроенные учетные записи
    $excludedAccounts = @(
        "BUILTIN\Administrators",
        "CREATOR OWNER", 
        "NT AUTHORITY\SYSTEM",
        "BUILTIN\Users",
        "NT AUTHORITY\Authenticated Users"
    )
    
    if ($excludedAccounts -contains $identity) {
        return $null
    }
    
    # Форматируем имя (убираем домен)
    if ($identity -like "*\*") {
        return $identity.Split('\')[1]
    }
    
    return $identity
}

# ----------------------------------------------------------------------------
# ЭТАП 1. Загружаем пользователей и группы из Active Directory
# ----------------------------------------------------------------------------
$usersInfo = Get-ADUsersInfo
$groupsInfo = Get-ADGroupsInfo

# Возвращает отображаемое имя (ФИО) по sAMAccountName.
# Если запись не найдена (например, уже удалённая учётная запись) —
# возвращается сам логин, чтобы строка не потерялась в отчёте.
function Get-DisplayName {
    param([string]$samAccountName)
    
    if ($usersInfo.ContainsKey($samAccountName)) {
        return $usersInfo[$samAccountName].DisplayName
    }
    elseif ($groupsInfo.ContainsKey($samAccountName)) {
        return $groupsInfo[$samAccountName].DisplayName
    }
    
    return $samAccountName
}

# Определяет, является ли учётная запись пользователем (а не группой):
# проверяется наличие логина в справочнике пользователей $usersInfo.
function Is-UserAccount {
    param([string]$samAccountName)
    
    return $usersInfo.ContainsKey($samAccountName)
}

# ----------------------------------------------------------------------------
# ЭТАП 2. Сканируем файловую систему и читаем ACL
# ----------------------------------------------------------------------------

# Показываем в консоли, какие пути не будут сканироваться.
Write-Host "Исключенные пути из сканирования:" -ForegroundColor Yellow
foreach ($excludedPath in $excludePaths) {
    Write-Host "  - $excludedPath" -ForegroundColor Yellow
}
Write-Host ""

# Для каждой папки читаем ACL (Get-Acl) и каждую запись доступа превращаем
# в строку отчёта: путь, учётная запись, ФИО, тип (пользователь/группа),
# тип доступа (разрешить/запретить), права и признак наследования.
Write-Host "Сканирование каталога $scanPath (максимальная глубина: 3 уровня)..." -ForegroundColor Green

# Собираем папки (глубина 3, с учётом $excludePaths).
# Сам корневой каталог добавляем отдельно — права, заданные на нём,
# тоже должны попасть в матрицу.
$allFolders = @()
if (Test-Path -LiteralPath $scanPath) {
    $allFolders += Get-Item -LiteralPath $scanPath
}
# @() защищает от добавления $null, если подпапок нет вовсе.
$allFolders += @(Get-FoldersRecursive -Path $scanPath -MaxDepth 3)
$allAccessData = @()   # Все записи ACL (и пользователи, и группы)
$groupAccessData = @() # Только записи групп — пригодятся на этапе 3

$progress = 0
foreach ($folder in $allFolders) {
    $progress++
    Write-Progress -Activity "Сканирование папок" -Status "Обработка: $($folder.FullName)" -PercentComplete (($progress / $allFolders.Count) * 100)
    
    try {
        $acl = Get-Acl -Path $folder.FullName -ErrorAction Stop
        
        foreach ($access in $acl.Access) {
            $formattedIdentity = Format-AccountName -identity $access.IdentityReference.Value
            
            if ($formattedIdentity) {
                $displayName = Get-DisplayName -samAccountName $formattedIdentity
                $translatedRights = Convert-RightsToRussian -rights $access.FileSystemRights.ToString()
                $accountType = if (Is-UserAccount -samAccountName $formattedIdentity) { "User" } else { "Group" }
                
                $accessRecord = [PSCustomObject]@{
                    Путь = $folder.FullName
                    УчетнаяЗапись = $formattedIdentity
                    ПолноеИмя = $displayName
                    ТипУчетнойЗаписи = $accountType
                    ТипДоступа = if ($access.AccessControlType -eq "Allow") { "Разрешить" } else { "Запретить" }
                    Права = $translatedRights
                    Наследование = if ($access.IsInherited) { "Наследуемый" } else { "Прямой" }
                }
                
                $allAccessData += $accessRecord
                
                if ($accountType -eq "Group") {
                    $groupAccessData += $accessRecord
                }
            }
        }
    }
    catch {
        Write-Warning "Ошибка доступа к папке: $($folder.FullName)"
    }
}

Write-Progress -Activity "Сканирование папок" -Completed

# ----------------------------------------------------------------------------
# ЭТАП 3. Разворачиваем доступы через группы
#   3.1. Сначала переносим прямые доступы пользователей.
#   3.2. Затем для каждого пользователя получаем все его группы (рекурсивно)
#        и добавляем доступы, которые эти группы дают на папки.
# ----------------------------------------------------------------------------
Write-Host "Формирование данных с учетом членства в группах..." -ForegroundColor Green

$userAccessData = @()
$userGroupMembership = @{}

# 3.1. Прямые доступы: пользователь явно указан в ACL папки
foreach ($access in $allAccessData) {
    if ($access.ТипУчетнойЗаписи -eq "User") {
        $userAccessData += [PSCustomObject]@{
            Путь = $access.Путь
            Пользователь = $access.УчетнаяЗапись
            ПолноеИмя = $access.ПолноеИмя
            ИсточникДоступа = "Прямой доступ"
            Группа = ""
            Права = $access.Права
            ТипДоступа = $access.ТипДоступа
            Наследование = $access.Наследование
        }
    }
}

# 3.2. Доступы через группы: для каждого пользователя разворачиваем членство
# и находим ACL-записи, где фигурируют его группы.
$userCount = 0
foreach ($userSamAccountName in $usersInfo.Keys) {
    $userCount++
    Write-Progress -Activity "Обработка пользователей" -Status "Пользователь: $userSamAccountName ($userCount из $($usersInfo.Count))" -PercentComplete (($userCount / $usersInfo.Count) * 100)
    
    # Получаем все группы пользователя (рекурсивно)
    $userGroups = Get-UserAllGroups -userSamAccountName $userSamAccountName -usersInfo $usersInfo -groupsInfo $groupsInfo
    $userGroupMembership[$userSamAccountName] = $userGroups
    
    # Для каждой группы пользователя находим соответствующие доступы
    foreach ($group in $userGroups) {
        $groupAccesses = $groupAccessData | Where-Object { $_.УчетнаяЗапись -eq $group }
        
        foreach ($groupAccess in $groupAccesses) {
            $userAccessData += [PSCustomObject]@{
                Путь = $groupAccess.Путь
                Пользователь = $userSamAccountName
                ПолноеИмя = $usersInfo[$userSamAccountName].DisplayName
                ИсточникДоступа = "Через группу"
                Группа = $group
                Права = $groupAccess.Права
                ТипДоступа = $groupAccess.ТипДоступа
                Наследование = $groupAccess.Наследование
            }
        }
    }
}

Write-Progress -Activity "Обработка пользователей" -Completed

# ----------------------------------------------------------------------------
# ЭТАП 4. Экспорт в CSV
# ----------------------------------------------------------------------------
Write-Host "Генерация CSV отчетов..." -ForegroundColor Green

# Основная матрица доступа: одна строка = пользователь + папка + права.
# Разделитель ";" выбран для корректного открытия в русской локали Excel.
$csvFilePath = Join-Path $outputDir "AccessMatrix_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$userAccessData | Export-Csv -Path $csvFilePath -Encoding UTF8 -NoTypeInformation -Delimiter ';'
Write-Host "CSV отчет сохранен: $csvFilePath" -ForegroundColor Green

# Дополнительный CSV: плоский список "пользователь → группа".
# Удобно для сверки и для построения сводных таблиц.
$membershipFilePath = Join-Path $outputDir "GroupMembership_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$membershipData = foreach ($user in $userGroupMembership.Keys) {
    foreach ($group in $userGroupMembership[$user]) {
        [PSCustomObject]@{
            Пользователь = $user
            ПолноеИмя = $usersInfo[$user].DisplayName
            Группа = $group
            ОтображаемоеИмяГруппы = Get-DisplayName -samAccountName $group
        }
    }
}
$membershipData | Export-Csv -Path $membershipFilePath -Encoding UTF8 -NoTypeInformation -Delimiter ';'
Write-Host "CSV с членством в группах сохранен: $membershipFilePath" -ForegroundColor Green

# Собираем уникальных пользователей для выпадающего списка
$uniqueUsers = $userAccessData | Select-Object Пользователь, ПолноеИмя | Sort-Object ПолноеИмя -Unique

# ----------------------------------------------------------------------------
# ЭТАП 5. Генерация интерактивного HTML-отчёта
# Отчёт — один HTML-файл: данные таблицы и список пользователей встраиваются
# при генерации, а фильтрация выполняется на JavaScript уже в браузере
# (интернет и сервер не нужны).
# ----------------------------------------------------------------------------
Write-Host "Генерация HTML отчета..." -ForegroundColor Green

$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Матрица доступа к файловым ресурсам</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; position: sticky; top: 0; }
        tr:hover { background-color: #f5f5f5; }
        .summary { background-color: #dff0d8; padding: 10px; margin-bottom: 15px; border-radius: 5px; }
        .filter-input { width: 100%; box-sizing: border-box; padding: 4px; margin-top: 5px; }
        .filter-header { background-color: #e9f7fe; }
        .controls { margin-bottom: 15px; }
        button { padding: 8px 15px; margin-right: 10px; margin-bottom: 5px; }
        .direct-access { background-color: #f0fff0; }
        .group-access { background-color: #fff0f0; }
        .info-badge { 
            background-color: #007bff; 
            color: white; 
            padding: 2px 6px; 
            border-radius: 3px; 
            font-size: 0.8em;
            margin-left: 5px;
        }
        .scan-info { 
            background-color: #e7f3ff; 
            padding: 8px; 
            margin-bottom: 10px; 
            border-radius: 4px;
            border-left: 4px solid #007bff;
        }
        .excluded-info { 
            background-color: #fff3cd; 
            padding: 8px; 
            margin-bottom: 10px; 
            border-radius: 4px;
            border-left: 4px solid #ffc107;
        }
    </style>
</head>
<body>
    <h1>Матрица доступа к файловым ресурсам</h1>
    
    <div class="scan-info">
        <strong>Область сканирования:</strong> $scanPath<br>
        <strong>Максимальная глубина:</strong> 3 уровня вложенности
    </div>
    
    <div class="excluded-info">
        <strong>Исключенные пути:</strong><br>
        <ul>
"@

# Добавляем информацию об исключенных путях в HTML отчет
foreach ($excludedPath in $excludePaths) {
    $htmlReport += "<li>$excludedPath</li>`n"
}

$htmlReport += @"
        </ul>
    </div>
    
    <div class="summary">
        <strong>Дата генерации:</strong> $(Get-Date -Format "dd.MM.yyyy HH:mm")<br>
        <strong>Всего записей:</strong> <span id="totalCount">$($userAccessData.Count)</span><br>
        <strong>Отфильтровано:</strong> <span id="filteredCount">$($userAccessData.Count)</span><br>
        <strong>Прямых доступов:</strong> $(($userAccessData | Where-Object { $_.ИсточникДоступа -eq "Прямой доступ" } | Measure-Object).Count)<br>
        <strong>Доступов через группы:</strong> $(($userAccessData | Where-Object { $_.ИсточникДоступа -eq "Через группу" } | Measure-Object).Count)<br>
        <strong>Обработано папок:</strong> $($allFolders.Count)
    </div>
    
    <div class="controls">
        <button onclick="clearFilters()">Очистить все фильтры</button>
        <button onclick="exportFilteredData()">Экспорт отфильтрованных данных</button>
        <button onclick="showAllData()">Показать все данные</button>
        <button onclick="showOnlyDirectAccess()">Только прямые доступы</button>
        <button onclick="showOnlyGroupAccess()">Только доступы через группы</button>
    </div>

    <table id="accessTable">
        <thead>
            <tr class="filter-header">
                <th>
                    Путь<br>
                    <input type="text" id="filterPath" class="filter-input" placeholder="Фильтр по пути..." onkeyup="filterTable()">
                </th>
                <th>
                    Пользователь<br>
                    <select id="filterUser" class="filter-input" onchange="filterTable()">
                        <option value="">Все пользователи</option>
"@

# Добавляем варианты для выпадающего списка (только пользователи)
foreach ($user in $uniqueUsers) {
    $displayText = if ($user.ПолноеИмя -and $user.ПолноеИмя -ne $user.Пользователь) {
        "$($user.ПолноеИмя) ($($user.Пользователь))"
    } else {
        $user.Пользователь
    }
    $htmlReport += "<option value=`"$($user.Пользователь)`">$displayText</option>`n"
}

$htmlReport += @"
                    </select>
                </th>
                <th>
                    Полное имя<br>
                    <input type="text" id="filterDisplayName" class="filter-input" placeholder="Фильтр по полному имени..." onkeyup="filterTable()">
                </th>
                <th>
                    Источник доступа<br>
                    <select id="filterSource" class="filter-input" onchange="filterTable()">
                        <option value="">Все источники</option>
                        <option value="Прямой доступ">Прямой доступ</option>
                        <option value="Через группу">Через группу</option>
                    </select>
                </th>
                <th>
                    Группа<br>
                    <input type="text" id="filterGroup" class="filter-input" placeholder="Фильтр по группе..." onkeyup="filterTable()">
                </th>
                <th>
                    Права<br>
                    <select id="filterRights" class="filter-input" onchange="filterTable()">
                        <option value="">Все права</option>
                        <option value="Полный доступ">Полный доступ</option>
                        <option value="Изменение">Изменение</option>
                        <option value="Чтение и выполнение">Чтение и выполнение</option>
                        <option value="Чтение">Чтение</option>
                        <option value="Запись">Запись</option>
                    </select>
                </th>
                <th>
                    Наследование<br>
                    <select id="filterInheritance" class="filter-input" onchange="filterTable()">
                        <option value="">Все</option>
                        <option value="Наследуемый">Наследуемый</option>
                        <option value="Прямой">Прямой</option>
                    </select>
                </th>
            </tr>
        </thead>
        <tbody id="tableBody">
"@

# Добавляем данные в таблицу
foreach ($item in $userAccessData) {
    $rowClass = if ($item.ИсточникДоступа -eq "Прямой доступ") { "direct-access" } else { "group-access" }
    $groupInfo = if ($item.ИсточникДоступа -eq "Через группу") { " ($($item.Группа))" } else { "" }
    
    $htmlReport += @"
            <tr class="$rowClass">
                <td>$($item.Путь)</td>
                <td>$($item.Пользователь)</td>
                <td>$($item.ПолноеИмя)</td>
                <td>$($item.ИсточникДоступа)$groupInfo</td>
                <td>$($item.Группа)</td>
                <td>$($item.Права)</td>
                <td>$($item.Наследование)</td>
            </tr>
"@
}

$htmlReport += @"
        </tbody>
    </table>

    <script>
        function filterTable() {
            const pathFilter = document.getElementById('filterPath').value.toLowerCase();
            const userFilter = document.getElementById('filterUser').value;
            const displayNameFilter = document.getElementById('filterDisplayName').value.toLowerCase();
            const sourceFilter = document.getElementById('filterSource').value;
            const groupFilter = document.getElementById('filterGroup').value.toLowerCase();
            const rightsFilter = document.getElementById('filterRights').value;
            const inheritanceFilter = document.getElementById('filterInheritance').value;
            
            const table = document.getElementById('accessTable');
            const rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
            let visibleCount = 0;
            
            for (let i = 0; i < rows.length; i++) {
                const cells = rows[i].getElementsByTagName('td');
                const path = cells[0].textContent.toLowerCase();
                const user = cells[1].textContent;
                const displayName = cells[2].textContent.toLowerCase();
                const source = cells[3].textContent;
                const group = cells[4].textContent.toLowerCase();
                const rights = cells[5].textContent;
                const inheritance = cells[6].textContent;
                
                const pathMatch = !pathFilter || path.includes(pathFilter);
                const userMatch = !userFilter || user === userFilter;
                const displayNameMatch = !displayNameFilter || displayName.includes(displayNameFilter);
                const sourceMatch = !sourceFilter || source.includes(sourceFilter);
                const groupMatch = !groupFilter || group.includes(groupFilter);
                const rightsMatch = !rightsFilter || rights.includes(rightsFilter);
                const inheritanceMatch = !inheritanceFilter || inheritance === inheritanceFilter;
                
                if (pathMatch && userMatch && displayNameMatch && sourceMatch && groupMatch && rightsMatch && inheritanceMatch) {
                    rows[i].style.display = '';
                    visibleCount++;
                } else {
                    rows[i].style.display = 'none';
                }
            }
            
            document.getElementById('filteredCount').textContent = visibleCount;
        }
        
        function clearFilters() {
            document.getElementById('filterPath').value = '';
            document.getElementById('filterUser').value = '';
            document.getElementById('filterDisplayName').value = '';
            document.getElementById('filterSource').value = '';
            document.getElementById('filterGroup').value = '';
            document.getElementById('filterRights').value = '';
            document.getElementById('filterInheritance').value = '';
            filterTable();
        }
        
        function showAllData() {
            clearFilters();
            filterTable();
        }
        
        function showOnlyDirectAccess() {
            clearFilters();
            document.getElementById('filterSource').value = 'Прямой доступ';
            filterTable();
        }
        
        function showOnlyGroupAccess() {
            clearFilters();
            document.getElementById('filterSource').value = 'Через группу';
            filterTable();
        }
        
        function exportFilteredData() {
            const table = document.getElementById('accessTable');
            const rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
            
            // Создаем CSV содержимое
            let csvContent = "Путь;Пользователь;Полное имя;Источник доступа;Группа;Права;Наследование\n";
            
            for (let i = 0; i < rows.length; i++) {
                if (rows[i].style.display !== 'none') {
                    const cells = rows[i].getElementsByTagName('td');
                    const row = [
                        cells[0].textContent,
                        cells[1].textContent,
                        cells[2].textContent,
                        cells[3].textContent,
                        cells[4].textContent,
                        cells[5].textContent,
                        cells[6].textContent
                    ];
                    csvContent += row.join(';') + '\n';
                }
            }
            
            // Создаем и скачиваем файл
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement("a");
            const url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", "filtered_access_matrix_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
        
        // Инициализация фильтрации при загрузке
        document.addEventListener('DOMContentLoaded', function() {
            filterTable();
        });
    </script>
</body>
</html>
"@

# Сохраняем HTML отчет
$htmlFilePath = Join-Path $outputDir "AccessMatrix_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$htmlReport | Out-File -FilePath $htmlFilePath -Encoding UTF8

Write-Host "HTML отчет сохранен: $htmlFilePath" -ForegroundColor Green

# ----------------------------------------------------------------------------
# ЭТАП 6. Генерируем вспомогательный скрипт SearchUserAccess.ps1
# Шаблон записан в одинарных кавычках (@' ... '@), чтобы PowerShell не
# подставлял переменные при генерации. Каталог отчётов подставляется
# отдельно через Replace() — так шаблон остаётся читаемым.
# ----------------------------------------------------------------------------
$userSearchScriptTemplate = @'
# Скрипт для поиска прав конкретного пользователя с учетом групп.
# ВНИМАНИЕ: файл создан автоматически скриптом FileAccessAudit.ps1
# и будет перезаписан при следующем запуске аудита.
param(
    [string]$UserName,
    [string]$ReportPath
)

if (-not $UserName) {
    $UserName = Read-Host "Введите имя пользователя для поиска"
}

# Каталог с отчётами аудита (подставлен при генерации).
$ReportDir = "__OUTPUT_DIR__"

# Если конкретный CSV не указан — берём самый свежий отчёт AccessMatrix_*.csv.
if (-not $ReportPath) {
    $latestReport = Get-ChildItem -Path $ReportDir -Filter "AccessMatrix_*.csv" |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1

    if (-not $latestReport) {
        Write-Host "Отчеты не найдены. Сначала запустите FileAccessAudit.ps1." -ForegroundColor Red
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
'@

# Подставляем реальный каталог отчётов вместо плейсхолдера __OUTPUT_DIR__.
$userSearchScript = $userSearchScriptTemplate.Replace('__OUTPUT_DIR__', $outputDir)

$userSearchScriptPath = Join-Path $scriptDir 'SearchUserAccess.ps1'
$userSearchScript | Out-File -FilePath $userSearchScriptPath -Encoding UTF8

Write-Host "Скрипт для поиска прав пользователя создан: $userSearchScriptPath" -ForegroundColor Green
Write-Host "Готово! Отчеты сохранены в: $outputDir" -ForegroundColor Green

# ----------------------------------------------------------------------------
# ЭТАП 7. Итоговая статистика в консоль
# ----------------------------------------------------------------------------
Write-Host "`nСтатистика:" -ForegroundColor Cyan
Write-Host "  Всего папок обработано: $($allFolders.Count)" -ForegroundColor Cyan
Write-Host "  Всего записей о доступах: $($userAccessData.Count)" -ForegroundColor Cyan
Write-Host "  Прямых доступов: $(($userAccessData | Where-Object { $_.ИсточникДоступа -eq 'Прямой доступ' } | Measure-Object).Count)" -ForegroundColor Cyan
Write-Host "  Доступов через группы: $(($userAccessData | Where-Object { $_.ИсточникДоступа -eq 'Через группу' } | Measure-Object).Count)" -ForegroundColor Cyan
Write-Host "  Уникальных пользователей: $(($uniqueUsers | Measure-Object).Count)" -ForegroundColor Cyan
Write-Host "  Исключенных путей: $($excludePaths.Count)" -ForegroundColor Cyan

Write-Host "`nИспользование:" -ForegroundColor Yellow
Write-Host "  Основной отчет: $htmlFilePath" -ForegroundColor Yellow
Write-Host "  Поиск прав пользователя: .\SearchUserAccess.ps1 -UserName 'имя_пользователя'" -ForegroundColor Yellow
