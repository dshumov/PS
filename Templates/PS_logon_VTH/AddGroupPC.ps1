cls
Import-Module ActiveDirectory

#ООО ВТХинвест
$Computers = Get-ADComputer -Filter * -SearchBase "OU=Компьютеры,OU=ООО ВТХинвест,DC=vth-invest,DC=ru"
foreach ($Computer in $Computers) {
	$Computer = $Computer.sAMAccountName
	add-adgroupmember "sg.vth.PC.All" -Members $Computer
	}
#ООО ТЖСК Тула
$Computers = Get-ADComputer -Filter * -SearchBase "OU=Компьютеры,OU=ООО ТЖСК Тула,DC=vth-invest,DC=ru"
foreach ($Computer in $Computers) {
	$Computer = $Computer.sAMAccountName
	add-adgroupmember "sg.tzsk.PC.All" -Members $Computer
	}
#АО Воздухотехника
$Computers = Get-ADComputer -Filter * -SearchBase "OU=Компьютеры,OU=АО Воздухотехника,DC=vth-invest,DC=ru"
foreach ($Computer in $Computers) {
	$Computer = $Computer.sAMAccountName
	add-adgroupmember "sg.zvth.PC.all" -Members $Computer
	}
#ООО Домодедово МУ-7
$Computers = Get-ADComputer -Filter * -SearchBase "OU=Компьютеры,OU=ООО Домодедово МУ-3,DC=vth-invest,DC=ru"
foreach ($Computer in $Computers) {
	$Computer = $Computer.sAMAccountName
	add-adgroupmember "sg.dmu.PC.all" -Members $Computer
	}
#ООО КСК
$Computers = Get-ADComputer -Filter * -SearchBase "OU=Компьютеры,OU=ООО КСК,DC=vth-invest,DC=ru"
foreach ($Computer in $Computers) {
	$Computer = $Computer.sAMAccountName
	add-adgroupmember "sg.ksk.PC.all" -Members $Computer
	}
#ООО ТД ТЖСК
$Computers = Get-ADComputer -Filter * -SearchBase "OU=Компьютеры,OU=ООО ТД ТЖСК,DC=vth-invest,DC=ru"
foreach ($Computer in $Computers) {
	$Computer = $Computer.sAMAccountName
	add-adgroupmember "sg.ttzsk.PC.All" -Members $Computer
	}
