Import-Module ActiveDirectory

#ООО ВТХинвест
$Dir = "\\vth-invest.ru\root\Персональные папки\ВТХинвест"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ВТХинвест,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}

#ООО ТЖСК Тула
$Dir = "\\vth-invest.ru\root\Персональные папки\ТЖСК"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТЖСК Тула,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}
	
#АО Воздухотехника
$Dir = "\\vth-invest.ru\root\Персональные папки\Воздухотехника"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=АО Воздухотехника,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}
	
#ООО КСК
$Dir = "\\vth-invest.ru\root\Персональные папки\КСК"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО КСК,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}
	
#ООО Домодедово МУ-3
$Dir = "\\vth-invest.ru\root\Персональные папки\МСУ"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО Домодедово МУ-3,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}
	
#ООО ТД Воздухотехника Мск
$Dir = "\\vth-invest.ru\root\Персональные папки\ТД Воздухотехника"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД Воздухотехника Мск,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}
	
#ООО ТД Воздухотехника СПб
$Dir = "\\vth-invest.ru\root\Персональные папки\ТД Воздухотехника"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД Воздухотехника СПб,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}
	
#ООО ТД ТЖСК
$Dir = "\\vth-invest.ru\root\Персональные папки\ТД ТЖСК"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД ТЖСК,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}

#ИП ВТХтелеком
$Dir = "\\vth-invest.ru\root\Персональные папки\ВТХтелеком"
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ИП ВТХтелеком,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
    $Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
    $ACL.SetAccessRule($Args)
    Set-Acl -Path $Path -AclObject $ACL
	Set-ADUser -Identity $User -HomeDrive "W:" -HomeDirectory "$Path1" -Server vsdc01.vth-invest.ru
  	}