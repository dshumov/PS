cls
#Подгрузка библиотек
Add-pssnapin Quest.ActiveRoles.ADManagement

#Файлы и программы
$Delete = "Начал удолять файлы."
$CompON = @()
$FileName = "Dell_Temp.ps1"
$Porsh = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$argument = "-executionpolicy unrestricted -file c:\$FileName"

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

		$comp
		$CompON += $comp
		
		Copy-Item "E:\_PowerShell\Компы\$FileName" -Destination \\$Comp\c$
		([wmiclass]"\\$Comp\root\cimv2:Win32_Process").create("$Porsh $argument")| Out-Null
		
	}
}

sleep -Seconds 10

$Delete

for ($i=0; $i -ne $CompON.Count; $i++)
{
	$Comp = $CompON[$i]
	Del "\\$Comp\c$\$FileName"
}
#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()