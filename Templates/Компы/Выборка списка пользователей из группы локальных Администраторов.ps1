cls

# Расположение файла Excel
$SetkaFail="E:\ГрАдм.xlsx"
# Имя листа Excel
$SheetName="Лист1"

$a = $null
$massGroup1 = @()
$massGroup2 = @()
$Complist = @()
$c = $null
$Comp = $null
$y = 0
$x = 0

# Получаем полный список компьютеров из каталога СР14
$Complist =  Get-QADComputer -SearchRoot "morton.ru/_Мортон/Ресурсы/Рабочие станции/" -IncludedProperties "name" -SizeLimit 5000 | 
	Where-Object {$_.Name -ne "CP777"}

# Перебор компов в массиве.
for ($p=0; $p -ne $Complist.Count; $p++)
{
	# Присваиваем переменной значения ячейки массива
	$Comp = $Complist[$p].name
	$Comp
	# Проверка активности компьютера.
	$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
	
	# Если Комп включен, выполняем копирование и запуск файла. 
	if ($c -eq "True")
	{
		$a = Get-WmiObject -Class Win32_GroupUser -ComputerName $Comp | Select-Object GroupComponent, PartComponent | 
		Where-Object {$_.GroupComponent -like "*Администраторы*" -and $_.PartComponent -notlike "*Administrator*" `
		-and $_.PartComponent -notlike "*admin*" -and $_.PartComponent -notlike "*Администраторы домена*" `
		-and $_.PartComponent -notlike "*sys.G.LocalAdmin*" -and $_.PartComponent -notlike "*sys.DL.cp14_PcAdmins*"} |

		ForEach-Object {$userGroup = $_.PartComponent -split '"', 0, "simplematch"; `
		$userGroup = $userGroup[3..3]; `
		$massGroup1 += $comp; `
		$massGroup2 += $userGroup}
	}
}


#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application
#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($SetkaFail)
#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName).UsedRange.Rows.Count

# Очищаем Excel файл от предыдущих значений суммы пользователей по организациям.
For ($Row=2;$Row -le $TotalsRow; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 1).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 2).Value()= ""
}

# Записываем массив с файлами в Excel файл.
$i = 0
For ($Row=2;$i -lt $massGroup1.Count; $Row++)
{

	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 1).Value()= $massGroup1[$i]
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 2).Value()= $massGroup2[$i]
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