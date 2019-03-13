cls
#Подгрузка библиотек
Add-pssnapin Quest.ActiveRoles.ADManagement

# Указываем путь в АД к каталогу где находятся УЗ юзеров.
$Catalog = "домен.ru/УЗ"
$UsersShara = "Где создаем новую папку пользователя"
$OffUsersShara = "Куда перемещаем папку удаленного пользователя"


$DirectoryName = $null

# Отправляем в конвеер включеных юзеров и отбираем тех у которых поле Факс пустое
Get-QADUser -SearchRoot $Catalog -Enabled | Where-Object {$_.Fax -eq $null} |
# Прописывает в поле Факс символ *
ForEach-Object {$_.Fax = *
# Составляет имя для личной папки пользователя
$DirectoryName = $_.DisplayName + "(" + $_.LogonName + ")"
# собираем конечный путь в директорию и создаем ее.
$UserDirectory = New-Item $UsersShara -Name $DirectoryName
# Помещаем настройки безопасности в переменую
$SecuritySeting = get-acl $UserDirectory
# Указывае пользователя для назначения прав
$identityReference = [System.Security.Principal.NTAccount] "$_.NTAccountName"
# Указываем права для пользователя
$fileSystemRights = [System.Security.AccessControl.FileSystemRights] 'FullControl'
# Указываем или ставим None нужно ли применять эти настройки к подпапкам и файлам.
$inheritanceFlags = [System.Security.AccessControl.InheritanceFlags] 'ContainerInherit,ObjectInherit'
# Указываем к кому применять ли свойства к самому каталогу и подкаталогам вложеных каталогов.
$propagationFlags = [System.Security.AccessControl.PropagationFlags] 'InheritOnly'
# Указываем разрешаем или запрещаем.
$accessControlType = [System.Security.AccessControl.accessControlType] 'Allow'
# Создаем объект настройки безопасности содержащий все наши настройки
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule($identityReference,
$fileSystemRights, $inheritanceFlags, $propagationFlags, $accessControlType)
# Добовляем наш объект к существующим настройкам.
$SecuritySeting.addAccessRule($rule)
# Записываем новые настройки безопасности к папке.
Set-Acl $UserDirectory -AclObject $SecuritySeting}

#
#
#Теперь расписываем перенос папок заблокированных пользователей.
# Отправляем в конвеер заблокированных юзеров и отбираем тех у которых поле Факс заполнено
Get-QADUser -SearchRoot $Catalog -Disabled | Where-Object {$_.Fax -eq $null} |
# Прописывает в поле Факс символ пустое значение
ForEach-Object {$_.Fax = $null
# Составляет имя как должна называться личная папка пользователя
$DirectoryName = $_.DisplayName + "(" + $_.LogonName + ")"
# собираем конечный путь в директорию
$UserDirectory = $UsersShara + "\" + $DirectoryName
# Перемещаем нашу папку в каталог для увольненых пользователей
Move-Item -Path $UserDirectory -Destination $OffUsersShara -Force
}