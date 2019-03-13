cls

# Расположение файла Excel
$SetkaFail="E:\ГрАдм.xlsx"
# Имя листа Excel
$SheetName="Лист3"

# Объявляем переменые
$Logs = @()
$CompList = @()
$UsersNames =@()
$Comp = ""
$UserName = ""
$i = 0
$sid = ""
$AdminName = ""
$admin = ""

#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application
#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($SetkaFail)
#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName).UsedRange.Rows.Count



For ($Row=2; $Row -lt $TotalsRow; $Row++)
{
	$AdminName = $objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 3).Value()
	
	if ($AdminName -ne "")
	{
		$CompList += @($objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 1).Value())
		$UsersNames += @($objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 2).Value())
	}
	else
	{
		$CompList += @("null")
		$UsersNames += @("null")
	}
	$AdminName = ""
}

$userList = Get-QADUser -SearchRoot "morton.ru/_Мортон/Учётные записи" -SizeLimit 7000 #| Where-Object {$_.sid -eq $Logs[$LN].ReplacementStrings[1]}

For($CR=0; $CR -lt $Complist.Count; $CR++)
{
	$Comp = $Complist[$CR]
	$Comp
	$UserName = $UsersNames[$CR]
	
	if($comp -ne "null")
	{
		# Проверка активности компьютера.
		$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
	 
		if ($c -eq "True")
		{
			$Logs = Get-EventLog -LogName Security -ComputerName $Comp | 
			Where-Object {$_.EventID -eq "4732"}
		
			if($Logs -ne $null)
			{
				for($LN=0; $LN -lt $Logs.Count; $LN++)
				{
					$sid = $Logs[$LN].ReplacementStrings[1]
					
					for ($i=0; $i -lt $userlist.count; $i++)
					{
						if($userlist[$i].sid -eq $sid)
						{
							$user = $userlist[$i].LogonN	
							if($user -eq $UserName)
							{
								$admin += @($Logs[$LN].ReplacementStrings[6])
								$date += @($Logs[$LN].TimeWritten)
								$LN = $Logs.Count
							}
							else
							{
								if($LN -eq $Logs.Count)
								{
									$admin += @("Журнал очищен")
									$date += @("Журнал очищен")
								}
							}						
							$i = $userlist.count
						}
					}				
				$sid = ""
				}
			}
			else
			{
				$admin += @("Журнал очищен")
				$date += @("Журнал очищен")
			}
		}
		else
		{
			$admin += @("null")
			$date += @("null")
		}
		$Logs = @()
		$user = ""
	}
	else
	{
		$admin += @("null")
		$date += @("null")
	}
}


$i = 0
For ($Row=2; $Row -lt $TotalsRow; $Row++)
{
	if($admin[$i] -ne "null")
	{
		$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 3).Value() = $admin[$i]
	}
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