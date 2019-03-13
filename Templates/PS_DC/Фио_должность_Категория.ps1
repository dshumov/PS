cls
#Подгрузка библиотек
Add-pssnapin Quest.ActiveRoles.ADManagement

# Расположение файла Excel
$SetkaFail="D:\Поле офис.xlsx"
# Имя листа Excel
$SheetName1="Лист2"

#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application

#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($SetkaFail)

#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName1).UsedRange.Rows.Count

# Очищаем Excel файл от предыдущих значений суммы пользователей по организациям.
For ($Row=2;$Row -le $TotalsRow; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName1).Cells.Item($Row, 1).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName1).Cells.Item($Row, 1).Value()= ""	
}

$Row = 2

$userlist =  Get-QADUser -SearchRoot "morton.ru/_Мортон/Учётные записи/СР14" -IncludedProperties "physicalDeliveryOfficeName" -Enabled

for ($NumberUser = 0; $NumberUser -le $userlist.Count; $NumberUser=$NumberUser + 1)
{
	$NumberUser
	$objWorkbook.Worksheets.Item($SheetName1).Cells.Item($Row, 1).Value() = $userlist[$NumberUser].Name
	$objWorkbook.Worksheets.Item($SheetName1).Cells.Item($Row, 2).Value() = $userlist[$NumberUser].Title
	$objWorkbook.Worksheets.Item($SheetName1).Cells.Item($Row, 3).Value() = $userlist[$NumberUser].Fax
	$Row++
}

#Сохраняем книгу Excel
$objExcel.Workbooks.save()
#Закрываем книгу Excel
$objExcel.Workbooks.close()
#Выходим из Excel (вернее даем команду на выход из Excel)
$objExcel.Quit()
#обнуляем объект
$objExcel = $null

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()