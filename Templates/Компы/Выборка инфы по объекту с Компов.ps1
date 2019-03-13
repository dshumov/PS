cls

# Расположение файла Excel
$SetkaFail="E:\Поляны.xlsx"
# Имя листа Excel
$SheetName="Лист1"

$Tema = "Поляны"
$Directory = ""
$mass = @()
$Complist = @()
$Comp = ""
$C = ""
$user = ""
$CompUser = ""
$massFiles = @()

New-Item E:\Слив\$Tema -ItemType Directory


if ($OfComp -eq $null)
{
	# Получаем полный список компьютеров из каталога СР14
	$Complist =  Get-QADComputer -SearchRoot "morton.ru/_Мортон/Ресурсы/Рабочие станции/СР14/" -IncludedProperties "name" | 
	Where-Object {$_.Name -ne "CP777" -and $_.Name -ne "CP555" -and $_.Name -ne "CP111"}
	
	# Перебор компов в массиве.
	for ($p=0; $p -ne $Complist.Count; $p++)
	{
		# Присваиваем переменной значения ячейки массива
		$Comp = $Complist[$p].name
		# Проверка активности компьютера.
		$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
	
		# Если Комп включен, выполняем копирование и запуск файла. 
		if ($c -eq "True")
		{
			$user = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName "$comp").UserName
			$user = $user -split "\", 0, "simplematch"
			$user = $user[1..1]
			
			$CompUser = "$comp" + "(" + "$user" + ")";
			
			$mass = @(gwmi -Class Win32_LogicalDisk -ComputerName $comp -Filter "DriveType='3'" | Where-Object {$_.Size -ne $null}) #| Select-Object DeviceID

			for ($i = 0; $i -ne $mass.Count; $i++)
			{
				$Device = $mass[$i].DeviceID -replace ":", "$"
				$DeviceName = $mass[$i].DeviceID -replace ":", ""
			
				Get-ChildItem "\\$comp\$Device\" | Where-Object {$_.Name -inotlike "*Windows*" `
				-and $_.Name -inotlike "*Temp*" -and $_.Name -inotlike "*Program Files*"} |
				ForEach-Object {Get-ChildItem $_.FullName -Filter "*$Tema*" -Recurse | 
				Where-Object {$_.PSIsContainer} |
				ForEach-Object {$massFiles += $_.FullName; `
				$Directory = $_.PSParentPath -split "$", 0, "simplematch"; `
				$Directory = $Directory[1..1]; `
				New-Item E:\Слив\$Tema\$CompUser\$DeviceName\$Directory\ -ItemType Directory; `
				Copy-Item $_.FullName E:\Слив\$Tema\$CompUser\$DeviceName\$Directory\ -Recurse -Force; `
				Remove-Item $_.FullName -Recurse -Force}}
			}						
		}
		else
		{
			$OfComp += @($Comp)
		}
		$Comp
		$Comp = ""
		$C = ""
	}
}
else
{
	$Complist = $OfComp
	$OfComp = @()
	
	for ($p=0; $p -ne $Complist.Count; $p++)
	{
		# Присваиваем переменной значения ячейки массива
		$Comp = $Complist[$p]
		# Проверка активности компьютера.
		$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
	
		# Если Комп включен, выполняем копирование и запуск файла. 
		if ($c -eq "True")
		{
			$user = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName "$comp").UserName
			$user = $user -split "\", 0, "simplematch"
			$user = $user[1..1]
			
			$CompUser = "$comp" + "(" + "$user" + ")"
			
			$mass = @(gwmi -Class Win32_LogicalDisk -ComputerName $comp -Filter "DriveType='3'" | Where-Object {$_.Size -ne $null}) #| Select-Object DeviceID

			for ($i = 0; $i -ne $mass.Count; $i++)
			{
				$Device = $mass[$i].DeviceID -replace ":", "$"
				$DeviceName = $mass[$i].DeviceID -replace ":", ""
				
				Get-ChildItem "\\$comp\$Device\" | Where-Object {$_.Name -inotlike "*Windows*" `
				-and $_.Name -inotlike "*Program Files (x86)*" -and $_.Name -inotlike "*Program Files*"} |
				ForEach-Object {Get-ChildItem $_.FullName -Filter "*$Tema*" -Recurse | 
				Where-Object {$_.PSIsContainer} |
				ForEach-Object {$massFiles += $_.FullName; `
				$Directory = $_.PSParentPath -split "$", 0, "simplematch"; `
				$Directory = $Directory[1..1]; `
				New-Item E:\Слив\$Tema\$CompUser\$DeviceName\$Directory\ -ItemType Directory; `
				Copy-Item $_.FullName E:\Слив\$Tema\$CompUser\$DeviceName\$Directory\ -Recurse -Force; `
				Remove-Item $_.FullName -Recurse -Force}}
			}	
		}
		else
		{
			$OfComp += @($Comp)
		}
		$Comp
		$Comp = ""
		$C = ""
	}	
}


#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application
#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($SetkaFail)
#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName).UsedRange.Rows.Count

# Очищаем Excel файл от предыдущих значений суммы пользователей по организациям.
For ($Row=1;$Row -le $TotalsRow; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 1).Value()= ""	
}

# Записываем массив с файлами в Excel файл.
$i = 0
For ($Row=1;$i -le $massFiles.Count; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 1).Value()= $massFiles[$i]
	$i++
}

#Закрываем книгу Excel
$objExcel.Workbooks.close()
#Выходим из Excel (вернее даем команду на выход из Excel)
$objExcel.Quit()
#обнуляем объект
$objExcel = $null


#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()