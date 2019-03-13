

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО КСК,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}