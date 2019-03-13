cls

# Имя запускаемого файла.
$FileName = "q.bat"

# Объявление переменных и массивов
$OfComp = @()
$Complist = @()
$Comp = ""
$C = ""

# Получаем полный список компьютеров из каталога СР14
$Complist =  Get-QADComputer -SearchRoot "morton.ru/_Мортон/Ресурсы/Рабочие станции/СР14/" -IncludedProperties "name"
# Перебор компов в массиве.
for ($i=0; $i -ne $Complist.Count; $i++)
{
	# Присваиваем переменной значения ячейки массива
	$Comp = $Complist[$i].name
	# Проверка активности компьютера.
	$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
	
	# Если Комп включен, выполняем копирование и запуск файла. 
	if ($c -eq "True")
	{
		Copy-Item G:\PowerGUI\Script\$FileName -Destination \\$Comp\c$\ -Force -Recurse
		([wmiclass]"\\$Comp\root\cimv2:Win32_Process").create("cmd.exe /c C:\$FileName") | Out-Null
		Start-Sleep -Seconds 2
		Set-Service -Name WinRM -ComputerName $Comp -StartupType Manual
		(Get-WmiObject win32_Service -Filter "name='WinRM'" -ComputerName $Comp).StopService() | Out-Null
		Remove-Item \\$Comp\c$\$FileName
		$Comp
	}
	else
	{
		$OfComp += $Comp
	}
	$Comp = ""
	$C = ""
}

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()