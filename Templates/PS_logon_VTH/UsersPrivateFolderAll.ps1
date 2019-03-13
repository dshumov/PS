cls
 
Import-Module ActiveDirectory
 
#АО Воздухотехника
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=АО Воздухотехника,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\Воздухотехника\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}

#ООО ВТХинвест
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ВТХинвест,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\ВТХинвест\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}
	
#ООО Домодедово МУ-3
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО Домодедово МУ-3,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\МСУ\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}
	
#ООО КСК
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО КСК,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\КСК\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}
	
#ООО ТД Воздухотехника Мск
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД Воздухотехника Мск,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\ТД Воздухотехника\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}
	
#ООО ТД Воздухотехника СПб
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД Воздухотехника СПб,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\ТД Воздухотехника\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}
	
#ООО ТД ТЖСК
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД ТЖСК,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\ТД ТЖСК\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}
	
#ООО ТЖСК Тула
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТЖСК Тула,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\ТЖСК\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}
	
#ИП ВТХтелеком
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ИП ВТХтелеком,DC=vth-invest,DC=ru" 
foreach ($User in $Users) {
	$User1 = $User.Name
	$User = $User.sAMAccountName
	$Dir = "\\vth-invest.ru\root\Персональные папки\ВТХтелеком\$User1"
	$Path = New-Item -ItemType Directory -Path $Dir -Name "Личная"
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"Modify, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $Args1 = New-Object  system.security.accesscontrol.filesystemaccessrule("VTH-INVEST\Администраторы домена","FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL = Get-Acl $Path
	$ACL.SetAccessRuleProtection($true,$false)
	$ACL.SetAccessRule($Args)
	$ACL.SetAccessRule($Args1)
	Set-Acl -Path $Path -AclObject $ACL
	}