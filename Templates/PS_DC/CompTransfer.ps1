cls
Import-Module ActiveDirectory

$Computers = Get-ADComputer -Filter * -SearchBase "CN=Computers,DC=vth-invest,DC=ru" 
foreach ($Computer in $Computers) {
	$Computer = $Computer.Name
    $id = $Computer.Remove(3)
		
	if ($id -eq "TZS") {Move-ADObject -Identity "CN=$Computer,CN=Computers,DC=vth-invest,DC=ru" -TargetPath "OU=Рабочие станции,OU=Компьютеры,OU=ООО ТЖСК Тула,DC=vth-invest,DC=ru"}
	elseif ($id -eq "VTH") {Move-ADObject -Identity "CN=$Computer,CN=Computers,DC=vth-invest,DC=ru" -TargetPath "OU=Рабочие станции,OU=Компьютеры,OU=ООО ВТХинвест,DC=vth-invest,DC=ru"}
	elseif ($id -eq "ZVT") {Move-ADObject -Identity "CN=$Computer,CN=Computers,DC=vth-invest,DC=ru" -TargetPath "OU=Рабочие станции,OU=Компьютеры,OU=АО Воздухотехника,DC=vth-invest,DC=ru"}
	elseif ($id -eq "TTZ") {Move-ADObject -Identity "CN=$Computer,CN=Computers,DC=vth-invest,DC=ru" -TargetPath "OU=Рабочие станции,OU=Компьютеры,OU=ООО ТД ТЖСК,DC=vth-invest,DC=ru"}
	elseif ($id -eq "KSK") {Move-ADObject -Identity "CN=$Computer,CN=Computers,DC=vth-invest,DC=ru" -TargetPath "OU=Рабочие станции,OU=Компьютеры,OU=ООО КСК,DC=vth-invest,DC=ru"}
	elseif ($id -eq "DMU") {Move-ADObject -Identity "CN=$Computer,CN=Computers,DC=vth-invest,DC=ru" -TargetPath "OU=Рабочие станции,OU=Компьютеры,OU=ООО Домодедово МУ-3,DC=vth-invest,DC=ru"}
	elseif ($id -eq "TEL") {Move-ADObject -Identity "CN=$Computer,CN=Computers,DC=vth-invest,DC=ru" -TargetPath "OU=Рабочие станции,OU=Компьютеры,OU=ИП ВТХтелеком,DC=vth-invest,DC=ru"}
	
  	}
	
	