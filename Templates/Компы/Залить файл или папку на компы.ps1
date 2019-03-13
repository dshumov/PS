cls


# Объявление переменных и массивов
$Complist = @()
$Comp = ""
$C = ""

if ($OfComp -eq $null)
{
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
			Copy-Item G:\ForUserPC\ExplodeProxy\dd-\dd-.VLX -Destination \\$Comp\D$\ExplodeProxy\dd- -Force -Recurse
			$Comp
		}
		else
		{
			$OfComp += $Comp
		}
		$Comp = ""
		$C = ""
	}
}
else
{
	$Complist = $OfComp
	$OfComp = @()
	
	for ($i=0; $i -ne $Complist.Count; $i++)
	{
		# Присваиваем переменной значения ячейки массива
		$Comp = $Complist[$i]
		# Проверка активности компьютера.
		$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
	
		# Если Комп включен, выполняем копирование и запуск файла. 
		if ($c -eq "True")
		{
			Copy-Item G:\ForUserPC\ExplodeProxy\dd-\dd-.VLX -Destination \\$Comp\D$\ExplodeProxy\dd- -Force -Recurse
			$Comp
		}
		else
		{
			$OfComp += $Comp
		}
		$Comp = ""
		$C = ""
	}	
}

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()