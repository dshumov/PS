<#
.SYNOPSIS
    Генерация HTML-телефонного справочника из Active Directory.

.DESCRIPTION
    Скрипт читает учётные записи пользователей из указанного OU, собирает ФИО,
    должность, отдел, телефоны, почту, руководителя и фотографию, формирует
    единый HTML-файл с фильтрами по фамилии и отделу и сохраняет его в общую
    папку. Дополнительно настраиваются NTFS-права на файл:
    чтение — пользователям домена, полный доступ — группе администраторов.

    HTML-файл самодостаточен: CSS и JavaScript встроены, фотографии
    закодированы в base64, внешние ресурсы не используются.

.NOTES
    Требования:
      * Windows PowerShell 5.1 или PowerShell 7+;
      * модуль ActiveDirectory (RSAT) и доступ к контроллеру домена;
      * права на запись в целевую папку и на изменение ACL файла.

    Справочник содержит персональные данные (ФИО, телефоны, почта, фото).
    Перед публикацией убедитесь, что доступ к файлу действительно ограничен.

.EXAMPLE
    .\phonebook.ps1
#>

# Модуль для работы с Active Directory (входит в набор RSAT).
Import-Module ActiveDirectory

# ============================================================================
#  НАСТРОЙКИ — заполните перед запуском
# ============================================================================

# Название организации — выводится в заголовке страницы.
# ЧТО ВСТАВИТЬ: название вашей компании.
$CompanyName = 'Моя компания'

# Куда сохранять готовый HTML-файл.
# ЧТО ВСТАВИТЬ: UNC-путь к общей папке на файловом сервере.
# Пример: '\\fileserver\share\phonebook.html'
$OutputPath = '\\fileserver\share\phonebook.html'

# OU, из которой берутся сотрудники.
# ЧТО ВСТАВИТЬ: distinguishedName вашего OU.
# Пример: 'OU=Users,OU=Company,DC=example,DC=local'
$SearchBase = 'OU=Users,OU=Company,DC=example,DC=local'

# Кому разрешено читать справочник.
# BUILTIN\Users обычно включает Domain Users.
# ЧТО ВСТАВИТЬ: 'BUILTIN\Users' или 'DOMAIN\Domain Users'.
# Пример: 'CONTOSO\Domain Users'
$ReadAccessGroup = 'BUILTIN\Users'

# NetBIOS-имя домена — используется при выдаче прав через icacls.
# ЧТО ВСТАВИТЬ: короткое имя домена.
# Пример: 'CONTOSO'
$DomainNetBIOS = 'CONTOSO'

# Группа, которой выдаётся полный доступ к файлу справочника.
# ЧТО ВСТАВИТЬ: имя группы администраторов файлового ресурса.
# Пример: 'FileShare Admins'
$FullAccessGroup = 'FileShare Admins'

# Получаем пользователей из AD.
# Набор свойств соответствует полям карточки сотрудника.
$Users = Get-ADUser -Filter * -SearchBase $SearchBase -Properties @(
    "cn", "mail", "title", "department", "company", 
    "telephonenumber", "mobile", "thumbnailphoto", 
    "jpegphoto", "employeeID", "info", "manager"
) | Sort-Object cn

# Список уникальных отделов — из него строится выпадающий фильтр.
$Departments = $Users | Where-Object { $_.department } | 
               Select-Object department -Unique | 
               Sort-Object department

# ----------------------------------------------------------------------------
# Шапка HTML: стили (CSS), заголовок и фильтры.
# Здесь начинается here-string, в который ниже добавляются варианты
# выпадающих списков (фамилии и отделы), карточки сотрудников и скрипт.
# ----------------------------------------------------------------------------
$HTMLHeader = @"
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Телефонный справочник $CompanyName</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .filters {
            background: #f8f9fa;
            padding: 20px;
            border-bottom: 1px solid #e9ecef;
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
            align-items: end;
        }
        
        .filter-group {
            flex: 1;
            min-width: 250px;
        }
        
        .filter-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
            color: #495057;
            font-size: 0.9em;
        }
        
        .filter-select {
            width: 100%;
            padding: 10px 15px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 14px;
            background: white;
            cursor: pointer;
            transition: all 0.3s ease;
        }
        
        .filter-select:focus {
            outline: none;
            border-color: #3498db;
            box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);
        }
        
        .users-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(400px, 1fr));
            gap: 20px;
            padding: 30px;
        }
        
        .user-card {
            background: white;
            border: 1px solid #e9ecef;
            border-radius: 12px;
            padding: 25px;
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        
        .user-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 15px 30px rgba(0,0,0,0.1);
            border-color: #3498db;
        }
        
        .user-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 5px;
            height: 100%;
            background: linear-gradient(135deg, #3498db, #2c3e50);
        }
        
        .user-header {
            display: flex;
            align-items: center;
            gap: 15px;
            margin-bottom: 20px;
        }
        
        .user-photo {
            width: 80px;
            height: 80px;
            border-radius: 50%;
            object-fit: cover;
            border: 3px solid #e9ecef;
            flex-shrink: 0;
        }
        
        .no-photo {
            width: 80px;
            height: 80px;
            border-radius: 50%;
            background: linear-gradient(135deg, #667eea, #764ba2);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 1.5em;
            border: 3px solid #e9ecef;
            flex-shrink: 0;
        }
        
        .user-name {
            font-size: 1.4em;
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 5px;
        }
        
        .user-title {
            color: #6c757d;
            font-size: 1em;
            font-weight: 500;
        }
        
        .user-details {
            display: grid;
            gap: 12px;
        }
        
        .detail-row {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            padding: 8px 0;
            border-bottom: 1px solid #f8f9fa;
        }
        
        .detail-label {
            font-weight: 600;
            color: #495057;
            min-width: 120px;
            font-size: 0.9em;
        }
        
        .detail-value {
            color: #2c3e50;
            flex: 1;
            font-size: 0.9em;
            line-height: 1.4;
        }
        
        .contact-info {
            background: #f8f9fa;
            padding: 12px;
            border-radius: 8px;
            margin-top: 5px;
        }
        
        .no-results {
            grid-column: 1 / -1;
            text-align: center;
            padding: 60px 20px;
            color: #6c757d;
            font-size: 1.2em;
        }
        
        .department-badge {
            display: inline-block;
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.8em;
            font-weight: 500;
            margin-top: 5px;
        }
        
        .reset-filters {
            background: #6c757d;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            transition: background 0.3s ease;
            margin-top: 5px;
        }
        
        .reset-filters:hover {
            background: #5a6268;
        }
        
        @media (max-width: 768px) {
            .users-grid {
                grid-template-columns: 1fr;
                padding: 15px;
            }
            
            .filters {
                flex-direction: column;
            }
            
            .filter-group {
                min-width: 100%;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📞 Телефонный справочник</h1>
            <p>$CompanyName — Все контакты сотрудников</p>
        </div>
        
        <div class="filters">
            <div class="filter-group">
                <label for="surnameFilter">👤 Фильтр по фамилии:</label>
                <select id="surnameFilter" class="filter-select" onchange="filterUsers()">
                    <option value="">Все сотрудники</option>
"@

# Варианты фильтра по фамилии: берём первое слово из cn (ФИО),
# убираем дубликаты и сортируем.
$Surnames = $Users | Where-Object { $_.cn } | ForEach-Object { 
    ($_.cn -split ' ')[0] 
} | Sort-Object | Get-Unique

$SurnameOptions = foreach ($Surname in $Surnames) {
    "<option value='$Surname'>$Surname</option>"
}

$HTMLHeader += $SurnameOptions -join "`n"

$HTMLHeader += @"
                </select>
            </div>
            <div class="filter-group">
                <label for="departmentFilter">🏢 Фильтр по отделу:</label>
                <select id="departmentFilter" class="filter-select" onchange="filterUsers()">
                    <option value="">Все отделы</option>
"@

# Варианты фильтра по отделу (только отделы, которые есть у сотрудников).
$DepartmentOptions = foreach ($Dept in $Departments) {
    if ($Dept.department) {
        "<option value='$($Dept.department)'>$($Dept.department)</option>"
    }
}

$HTMLHeader += $DepartmentOptions -join "`n"

$HTMLHeader += @"
                </select>
            </div>
            <div class="filter-group">
                <button class="reset-filters" onclick="resetFilters()">🔄 Сбросить фильтры</button>
            </div>
        </div>
        
        <div class="users-grid" id="usersGrid">
"@

# Подвал HTML и JavaScript фильтрации.
# Фильтры работают в браузере по атрибутам data-surname и data-department
# карточек — перезагрузка страницы и сервер не нужны.
$HTMLFooter = @"
        </div>
    </div>

    <script>
        function filterUsers() {
            const surnameFilter = document.getElementById('surnameFilter').value.toLowerCase();
            const departmentFilter = document.getElementById('departmentFilter').value.toLowerCase();
            const userCards = document.querySelectorAll('.user-card');
            let visibleCount = 0;
            
            userCards.forEach(card => {
                const surname = card.getAttribute('data-surname').toLowerCase();
                const department = card.getAttribute('data-department').toLowerCase();
                
                const surnameMatch = !surnameFilter || surname === surnameFilter;
                const departmentMatch = !departmentFilter || department === departmentFilter;
                
                if (surnameMatch && departmentMatch) {
                    card.style.display = 'block';
                    visibleCount++;
                } else {
                    card.style.display = 'none';
                }
            });
            
            // Show/hide no results message
            let noResults = document.getElementById('noResults');
            if (visibleCount === 0) {
                if (!noResults) {
                    noResults = document.createElement('div');
                    noResults.id = 'noResults';
                    noResults.className = 'no-results';
                    noResults.innerHTML = '😔 По вашему запросу ничего не найдено';
                    document.getElementById('usersGrid').appendChild(noResults);
                }
            } else if (noResults) {
                noResults.remove();
            }
        }
        
        function resetFilters() {
            document.getElementById('surnameFilter').value = '';
            document.getElementById('departmentFilter').value = '';
            filterUsers();
        }
        
        // Initial filter on page load
        document.addEventListener('DOMContentLoaded', function() {
            filterUsers();
        });
    </script>
</body>
</html>
"@

# Функция кодирует фотографию из AD (байты) в data:URL base64,
# чтобы вставить её прямо в HTML и не зависеть от внешних файлов.
# Если фото в AD нет — возвращает $null, и в карточке рисуется заглушка.
function Convert-PhotoToBase64 {
    param($PhotoData)
    if ($PhotoData) {
        $Base64 = [Convert]::ToBase64String($PhotoData)
        return "data:image/jpeg;base64,$Base64"
    }
    return $null
}

# Получает ФИО руководителя по DN из атрибута manager.
# Если пользователь не найден — возвращает CN из DN или «Не указан».
function Get-ManagerName {
    param($ManagerDN)
    if ($ManagerDN) {
        try {
            $Manager = Get-ADUser -Identity $ManagerDN -Properties "DisplayName" -ErrorAction SilentlyContinue
            if ($Manager) {
                return $Manager.DisplayName
            }
        }
        catch {
            # If cannot get manager details, return the CN from DN
            if ($ManagerDN -match "CN=([^,]+)") {
                return $Matches[1]
            }
        }
    }
    return "Не указан"
}

# Формируем карточку каждого сотрудника.
# В карточку попадают только заполненные поля (проверки if внутри here-string).
$UserCardsHTML = @()
foreach ($User in $Users) {
    # Пропускаем записи без ФИО (служебные учётные записи и т. п.).
    if (-not $User.cn) { continue }
    
    # Фамилия = первое слово из cn; используется фильтром по фамилии.
    $Surname = ($User.cn -split ' ')[0]
    
    # Фото: сначала thumbnailPhoto (стандартный атрибут), если пусто — jpegPhoto.
    $UserPhoto = Convert-PhotoToBase64 -PhotoData $User.thumbnailphoto
    if (-not $UserPhoto) {
        $UserPhoto = Convert-PhotoToBase64 -PhotoData $User.jpegphoto
    }
    
    # Get manager name
    $ManagerName = Get-ManagerName -ManagerDN $User.manager
    
    # Build user card HTML
    $UserCard = @"
            <div class="user-card" data-surname="$Surname" data-department="$($User.department)">
                <div class="user-header">
                    $(if ($UserPhoto) {
                        "<img src='$UserPhoto' alt='Фото $($User.cn)' class='user-photo'>"
                    } else {
                        "<div class='no-photo'>$($User.cn[0])</div>"
                    })
                    <div>
                        <div class="user-name">$($User.cn)</div>
                        <div class="user-title">$($User.title)</div>
                        $(if ($User.department) {
                            "<div class='department-badge'>$($User.department)</div>"
                        })
                    </div>
                </div>
                
                <div class="user-details">
                    $(if ($User.mail) {
                        "<div class='detail-row'>
                            <span class='detail-label'>📧 Email:</span>
                            <span class='detail-value'><a href='mailto:$($User.mail)'>$($User.mail)</a></span>
                        </div>"
                    })
                    
                    $(if ($User.company) {
                        "<div class='detail-row'>
                            <span class='detail-label'>🏢 Организация:</span>
                            <span class='detail-value'>$($User.company)</span>
                        </div>"
                    })
                    
                    $(if ($User.department) {
                        "<div class='detail-row'>
                            <span class='detail-label'>📊 Отдел:</span>
                            <span class='detail-value'>$($User.department)</span>
                        </div>"
                    })
                    
                    $(if ($User.telephonenumber -or $User.mobile) {
                        "<div class='detail-row'>
                            <span class='detail-label'>📞 Телефоны:</span>
                            <span class='detail-value'>
                                <div class='contact-info'>
                                    $(if ($User.telephonenumber) {
                                        "📞 Рабочий: $($User.telephonenumber)<br>"
                                    })
                                    $(if ($User.mobile) {
                                        "📱 Мобильный: $($User.mobile)"
                                    })
                                </div>
                            </span>
                        </div>"
                    })
                    
                    $(if ($User.employeeID) {
                        "<div class='detail-row'>
                            <span class='detail-label'>🎂 День рождения:</span>
                            <span class='detail-value'>$($User.employeeID)</span>
                        </div>"
                    })
                    
                    $(if ($ManagerName) {
                        "<div class='detail-row'>
                            <span class='detail-label'>👤 Руководитель:</span>
                            <span class='detail-value'>$ManagerName</span>
                        </div>"
                    })
                    
                    $(if ($User.info) {
                        "<div class='detail-row'>
                            <span class='detail-label'>📝 Заметки:</span>
                            <span class='detail-value'>$($User.info)</span>
                        </div>"
                    })
                </div>
            </div>
"@
    $UserCardsHTML += $UserCard
}

# Склеиваем страницу: шапка + все карточки + подвал.
$FullHTML = $HTMLHeader + ($UserCardsHTML -join "`n") + $HTMLFooter

# Сохраняем готовый файл. -Force перезаписывает предыдущую версию.
# Кодировка UTF8 (с BOM) — корректная кириллица в браузере.
try {
    $FullHTML | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
    Write-Host "✅ HTML файл успешно создан: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Host "❌ Ошибка при сохранении файла: $_" -ForegroundColor Red
    exit 1
}

# ----------------------------------------------------------------------------
# NTFS-права на готовый файл:
#   * чтение — группа $ReadAccessGroup (по умолчанию BUILTIN\Users);
#   * полный доступ — группа администраторов $FullAccessGroup.
# Права выставляются утилитой icacls; /inheritance:r отключает наследование,
# чтобы доступ к справочнику не зависел от прав родительской папки.
# ----------------------------------------------------------------------------
try {
    Write-Host "🔧 Проверка групп и установка прав доступа..." -ForegroundColor Cyan

    $FullAccessGroupExists = $false
    $FullAccessGroupName = $null

    # Ищем группу полного доступа: сначала заданное имя, затем типовые варианты.
    $CandidateNames = @()
    if (-not [string]::IsNullOrWhiteSpace($FullAccessGroup)) {
        $CandidateNames += $FullAccessGroup
    }
    $CandidateNames += @(
        'Full Access to Resources',
        'FullAccess',
        'Resource Admins',
        'File Share Admins'
    )

    foreach ($CandidateName in $CandidateNames) {
        try {
            $Group = Get-ADGroup -Filter "Name -like '*$CandidateName*'" -ErrorAction Stop |
                     Select-Object -First 1
        }
        catch {
            $Group = $null
        }

        if ($Group) {
            $FullAccessGroupExists = $true
            $FullAccessGroupName = $Group.Name
            Write-Host "✅ Найдена группа: $($FullAccessGroupName)" -ForegroundColor Green
            break
        }
    }

    if (-not $FullAccessGroupExists) {
        Write-Host "⚠️ Группа '$FullAccessGroup' не найдена — используется запасной вариант." -ForegroundColor Yellow
    }

    # /reset снимает явные ACE и возвращает наследование от родительской папки.
    Write-Host "🔄 Сброс прав доступа..." -ForegroundColor Cyan
    icacls $OutputPath /reset

    # Чтение для пользователей домена.
    Write-Host "🔐 Установка прав на чтение: $ReadAccessGroup..." -ForegroundColor Cyan
    icacls $OutputPath /grant:r "${ReadAccessGroup}:(R)"

    if ($FullAccessGroupExists) {
        Write-Host "🔐 Установка прав для группы '$FullAccessGroupName'..." -ForegroundColor Cyan

        # Вариант 1: домен\группа.
        icacls $OutputPath /grant:r "${DomainNetBIOS}\${FullAccessGroupName}:(F)" /inheritance:r

        if ($LASTEXITCODE -ne 0) {
            # Вариант 2: имя группы без домена.
            icacls $OutputPath /grant:r "${FullAccessGroupName}:(F)" /inheritance:r
        }

        if ($LASTEXITCODE -ne 0) {
            # Вариант 3: через SID — если имя группы не разрешается icacls.
            $GroupSID = (Get-ADGroup -Identity $FullAccessGroupName).SID.Value
            icacls $OutputPath /grant:r "*${GroupSID}:(F)" /inheritance:r
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ Права для '$FullAccessGroupName' установлены" -ForegroundColor Green
        }
        else {
            Write-Host "❌ Не удалось установить права для группы '$FullAccessGroupName'" -ForegroundColor Red
        }
    }
    else {
        # Запасной вариант: полный доступ только администраторам.
        icacls $OutputPath /grant:r 'BUILTIN\Administrators:(F)' /inheritance:r
        Write-Host "✅ Права для BUILTIN\Administrators установлены" -ForegroundColor Green
    }

    Write-Host "✅ Процесс установки прав завершен" -ForegroundColor Green
}
catch {
    Write-Host "❌ Критическая ошибка при установке прав: $_" -ForegroundColor Red
    Write-Host "💡 Рекомендации:" -ForegroundColor Yellow
    Write-Host "   1. Проверьте, что группа '$FullAccessGroup' существует в AD" -ForegroundColor White
    Write-Host "   2. Убедитесь, что у вас есть права на изменение ACL" -ForegroundColor White
    Write-Host "   3. Установите права вручную командой:" -ForegroundColor White
    Write-Host "      icacls `"$OutputPath`" /reset /grant:r `"${ReadAccessGroup}:(R)`" /grant:r `"${DomainNetBIOS}\${FullAccessGroup}:(F)`" /inheritance:r" -ForegroundColor White
}

# Открываем готовый справочник в браузере по умолчанию.
# Если открыть не удалось (например, нет браузера в сессии) — просто
# показываем путь, по которому лежит файл.
try {
    Start-Process $OutputPath
    Write-Host "🌐 Файл открывается в браузере..." -ForegroundColor Cyan
}
catch {
    Write-Host "⚠️ Не удалось открыть файл в браузере: $_" -ForegroundColor Yellow
    Write-Host "📁 Файл сохранен по пути: $OutputPath" -ForegroundColor Cyan
}
