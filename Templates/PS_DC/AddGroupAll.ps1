cls
Import-Module ActiveDirectory

#ООО ВТХинвест
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ВТХинвест,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.vth.All" -Members $User
	add-adgroupmember "ВТХинвест" -Members $User
	}
#ООО ТЖСК Тула
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТЖСК Тула,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.tzsk.All" -Members $User
	add-adgroupmember "ТЖСК" -Members $User
	}
#АО Воздухотехника
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=АО Воздухотехника,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.zvth.all" -Members $User
	add-adgroupmember "АО Воздухотехника" -Members $User
	}
#ООО Домодедово МУ-7
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО Домодедово МУ-3,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.dmu.all" -Members $User
	add-adgroupmember "ДМУ-7" -Members $User
	}
#ООО КСК
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО КСК,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.ksk.all" -Members $User
	add-adgroupmember "КСК" -Members $User
	}
#ООО ТД ТЖСК
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО ТД ТЖСК,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.ttzsk.All" -Members $User
	add-adgroupmember "ТД ТЖСК" -Members $User
	}
#ИП ВТХтелеком
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ИП ВТХтелеком,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.tel.All" -Members $User
	add-adgroupmember "ВТХтелеком" -Members $User
	}
#ООО Тридика
$Users = Get-ADUser -Filter * -SearchBase "OU=Пользователи,OU=ООО Тридика,DC=vth-invest,DC=ru"
foreach ($User in $Users) {
	$User = $User.sAMAccountName
	add-adgroupmember "sg.vsfs01.root.all.L" -Members $User
	add-adgroupmember "sg.trd.All" -Members $User
	add-adgroupmember "Тридика" -Members $User
	}