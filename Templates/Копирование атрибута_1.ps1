$users = Get-ADUser  -Filter * -searchBase "OU=Филиалы,OU=Users,OU=MSK,DC=vostok-electra,DC=ru" -Properties *

foreach ($user in $users)
{
$branch = $user.Department
SET-ADUSER $user –replace @{branch=$branch}
SET-ADUSER $user –replace @{Company="Восток"}

}  
  
