# Путь к OU пользователей
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=АО Воздухотехника,DC=vth-invest,DC=ru" 
 

foreach ($User in $Users) {
$User = $User.sAMAccountName
Disable-Mailbox -Identity $User
    
	}