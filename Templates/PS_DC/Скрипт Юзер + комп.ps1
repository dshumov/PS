cls
#Подгрузка библиотек
Add-pssnapin Quest.ActiveRoles.ADManagement

# Расположение файла Excel
$SetkaFail="E:\UsersComp.xlsx"
# Имя листа Excel
$SheetName="Лист1"
									# Переменные массивов

$MassUser = @()
$MassKomp1 = @()
$MassKomp2 = @()
$MassKomp3 = @()



# Сканируем АД на пользователей отвечающих фильтрам.
$userlist =  Get-QADUser -SearchRoot "morton.ru/_Мортон/Учётные записи/СР14/" -IncludedProperties "info" -MemberOf "agg.G.eToken" -Email "*@*" -Enabled -SizeLimit 7000


													# Начинаем обработку данных из Excel и данных из АД


# Для пользователя из списка перебором.
for ($NumberUser = 0; $NumberUser -lt $userlist.Count; $NumberUser=$NumberUser + 1)
{	
	# Выводим номер обрабатываемого пользователя
	$NumberUser
	
	# Если поле "Info" не пустое выполняем.
	if ($userlist[$NumberUser].info -ne $null)
	{
		#Записываем юзера в массив пользователей.
		$MassUser += $userlist[$NumberUser].Name
		
		# Разбаваем поле info на массив и пересобираем массив оставляя только имена компьютеров.
		$info = -split $userlist[$NumberUser].info
		
		# Если переменная с именем компьютера1 не пустая
		if ($info[0..0] -ne "")
		{
			$MassKomp1 += $info[0..0]
		}
		else
		{
			$MassKomp1 += '-'
		}
		
		# Если переменная с именем компьютера2 не пустая
		if ($info[3..3] -ne "")
		{
			$MassKomp2 += $info[3..3]
		}
		else
		{
			$MassKomp2 += '-'
		}
		
		# Если переменная с именем компьютера3 не пустая
		if ($info[6..6] -ne "")
		{
			$MassKomp3 += $info[6..6]
		}
		else
		{
			$MassKomp3 += '-'
		}
	}
}	
																	#Записываем все масивы в файл Excel.
																		
#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application


#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($SetkaFail)																		

																		#ПРОИЗВОДИМ ОЧИСТКУ ФАЙЛА Excel.
													
#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName).UsedRange.Rows.Count

# Очищаем Excel файл от предыдущих значений суммы пользователей по организациям.
For ($Row=2;$Row -le $TotalsRow; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 1).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 2).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 3).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($Row, 4).Value()= ""
}

$p = 2

# Записываем лист "Подсчет" Имена офисов и счетчик.
For ($i = 0; $i -le $MassUser.Count; $i++)
{
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($p, 1).Value() = $MassUser[$i]
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($p, 2).Value() = $MassKomp1[$i]
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($p, 3).Value() = $MassKomp2[$i]
	$objWorkbook.Worksheets.Item($SheetName).Cells.Item($p, 4).Value() = $MassKomp3[$i]
	$p++
}
$p = 2


#Закрываем книгу Excel
$objExcel.Workbooks.close()
#Выходим из Excel (вернее даем команду на выход из Excel)
$objExcel.Quit()
#обнуляем объект
$objExcel = $null

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()