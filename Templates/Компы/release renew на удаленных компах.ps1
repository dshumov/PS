cls

#Файлы и программы

$Porsh = "cmd.exe" 
$argument1 = "ipconfig /rlease"
$argument2 = "ipconfig /renew"
$argument2 = "gpupdate /force"


	# Присваиваем переменной значения ячейки массива
	$Comp = "cp777"
	# Проверка активности компьютера.
	$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
	
	# Если Комп включен, выполняем копирование и запуск файла. 
	if ($c -eq "True")
	{
		$comp
		([wmiclass]"\\$Comp\root\cimv2:Win32_Process").create("$Porsh /c $argument1")
		sleep -Seconds 1
		([wmiclass]"\\$Comp\root\cimv2:Win32_Process").create("$Porsh /c $argument2") | Out-Null
		sleep -Seconds 1
		([wmiclass]"\\$Comp\root\cimv2:Win32_Process").create("$Porsh /c $argument3") | Out-Null
	}

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()