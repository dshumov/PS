#ООО ТД ТЖСК

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД ТЖСК,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}