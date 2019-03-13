cls
#Подгрузка библиотек
Add-pssnapin Quest.ActiveRoles.ADManagement

# Указываем путь в АД к каталогу где находятся УЗ юзеров.
$Catalog = "домен.ru/УЗ"

# Отправляем в конвеер юзеров и изменяем имя УЗ на формат Фамилия+Имя.
Get-QADUser -SearchRoot $Catalog -Enabled | 
ForEach-Object {$_.name = $_.LastName + " " + $_.givenName
$_.DisplayName = $_.LastName + " " + $_.givenName}

