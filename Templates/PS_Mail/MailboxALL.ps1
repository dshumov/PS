#АО Воздухотехника

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=АО Воздухотехника,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}

#ООО ВТХинвест

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ВТХинвест,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}

#ООО Домодедово МУ-3

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО Домодедово МУ-3,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}

#ООО КСК

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО КСК,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}

#ООО ТД Воздухотехника Мск

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД Воздухотехника Мск,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}

#ООО ТД Воздухотехника СПб

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД Воздухотехника СПб,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}

#ООО ТД ТЖСК

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД ТЖСК,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}

#ООО ТЖСК Тула

$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТЖСК Тула,DC=vth-invest,DC=ru"

foreach ($User in $Users) {
$User = $User.sAMAccountName

Enable-Mailbox -Identity $User

}