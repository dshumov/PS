cls
#Подгрузка библиотек
Add-pssnapin Quest.ActiveRoles.ADManagement

# Расположение файла Excel
$SetkaFail="D:\Сеть.xlsx"
# Имя листа Excel
$SheetName1="подсети"
$SheetName2="Подсчет"
$SheetName3="НеопПодсети"
$SheetName4="КолКомпов"

									# Переменные массивов
$Mas = @()
$Mass = @()
$MassExcelOutOffice = @()
$MassExcelOutIP = @()
$MassUserOffice = @()
$MassUserIP = @()
$MassUserSumOffice = @()
$MassUserSumIP = @()
$MassKompOffice = @()
$MassKompCompany = @()
$MassKompName = @()
$MassKompSum = @()
$MassKOffice = @()
$MassKSum = @()

									# Объявляем переменные.
									# Числовые переменные
$n = 0
$i = 0
$KompRepeat = 0
$KompUnknown = 0
									# Символьние переменные
$NameKomp = ""
$Number = ""

					# Сначала забираем все данные о подсетях из файла Excel

#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application


#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($SetkaFail)

													#ПРОИЗВОДИМ ОЧИСТКУ ФАЙЛА Excel.
													
#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName2).UsedRange.Rows.Count

# Очищаем Excel файл от предыдущих значений суммы пользователей по организациям.
For ($Row=2;$Row -le $TotalsRow; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName2).Cells.Item($Row, 1).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName2).Cells.Item($Row, 2).Value()= ""	
}

#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName3).UsedRange.Rows.Count

# Очищаем Excel файл от предыдущих значений суммы пользователей по организациям.
For ($Row=2;$Row -le $TotalsRow; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName3).Cells.Item($Row, 1).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName3).Cells.Item($Row, 2).Value()= ""	
}

#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName4).UsedRange.Rows.Count

# Очищаем Excel файл от предыдущих значений суммы пользователей по организациям.
For ($Row=2;$Row -le $TotalsRow; $Row++)
{
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($Row, 1).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($Row, 2).Value()= ""
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($Row, 3).Value()= ""
}

$objWorkbook.Worksheets.Item($SheetName4).Cells.Item(2, 5).Value()= ""
$objWorkbook.Worksheets.Item($SheetName4).Cells.Item(2, 6).Value()= ""

													#ПРОИЗВОДИМ ВЫБОРКУ ДАННЫХ ИЗ ФАЙЛА Excel.
													
#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName1).UsedRange.Rows.Count

# Перебором добавляем в массивы значения поля офис и поля подсеть относительно друг друга.
For ($Row=2;$Row -le $TotalsRow; $Row++)
{
	$MassExcelOutOffice += $objWorkbook.Worksheets.Item($SheetName1).Cells.Item($Row, 1).Value()
	$ip = $objWorkbook.Worksheets.Item($SheetName1).Cells.Item($Row, 3).Value()
	
	if ($ip -ne "")
	{
		$ip = $ip -split ".", 0, "simplematch"
		$ip = $ip[0..2]
		$ip = $ip -join "."
	}
	$MassExcelOutIP += $ip
}


# Ванлайнер для поиска в массиве номера элемента массива искомого нами значение.
function findinarr ($array, $value) {for ($i=0; $i -lt $array.count;$i++){if($array[$i] -eq $value){$i}}}

# Сканируем АД на пользователей отвечающих фильтрам.
$userlist =  Get-QADUser -SearchRoot "morton.ru/_Мортон/Учётные записи/" -IncludedProperties "info" -MemberOf "agg.G.eToken" -Email "*@*" -Enabled -SizeLimit 7000


													# Начинаем обработку данных из Excel и данных из АД


# Для пользователя из списка перебором.
for ($NumberUser = 0; $NumberUser -lt $userlist.Count; $NumberUser=$NumberUser + 1)
{	
	# Выводим номер обрабатываемого пользователя
	$NumberUser
	
	# Если поле "Info" не пустое выполняем.
	if ($userlist[$NumberUser].info -ne $null)
	{
		# Разбаваем поле info на массив и пересобираем массив оставляя только имена компьютеров.
		$info = -split $userlist[$NumberUser].info
		$info = $info[0..0] + $info[3..3] + $info[6..6]
		
		# Пока счетчек перебираемых компов из массива меньше или равен 2
		while ($n -le 2)
		{
			# Если имя компьютера в массиве начинается на "VS", переходим к следующему элементу массива.
			# Иначе записываем имя компа в переменую, приводя переменную счетчика к значению для завершения цикла.
			if ( $info[$n] -like "VS*")
			{
				$n++
			}
			else
			{
				$NameKomp = $info[$n]
				$n = 3
			}
		}
		$n = 0
		
		# Если переменная с именем компьютера не пустая
		if ($NameKomp -ne "")
		{
			# Присваеваем переменной "ip" из DNS значение IP, компьютера полученого из поля "info".
			$ip = [System.Net.Dns]::GetHostAddresses($NameKomp)[0].ipaddresstostring
			
		
			# Проверяем что IP адрес был получен из DNS.
			if ($ip -ne "")
			{
				# Приводим полученый IP к виду до подсети.
				$ip = $ip -split ".", 0, "simplematch"
				$ip = $ip[0..2]
				$ip = $ip -join "."
				
				
				# Если полученая подсеть входит в масив полученый из файла Excel
				if ($MassExcelOutIP -eq $ip)		########################################################################################
				{
					# Находим каким элементом массива является наш IP
					$Number = findinarr $MassExcelOutIP $ip
					$OfficeName = $MassExcelOutOffice[$Number]
					
																			# ПОДСЧЕТ ЮЗЕРОВ
					
					# Если в заполняемом масиве с полем офис для подсчета юзеров уже записано найденое поле офис.
					if ($MassUserOffice -eq $OfficeName)
					{
						# Получаем индекс этого элемента офис и увеличиваем значение счетчика в масиве с подсчетом.
						$a = findinarr $MassUserOffice $OfficeName
						$b = $MassUserSumOffice[$a]
						$b++
						$MassUserSumOffice[$a] = $b
					}
					else
					{
						# Новые значения для счетчика в массив.
						$MassUserOffice += $OfficeName
						$MassUserSumOffice += 1
					}
					
																			# ПОДСЧЕТ КОМПОВ
					
					# Проверяем наличие имени компьютера в массиве со всеми обработанами из файла Excel именами компьютеров.
					if($MassKompName -eq $NameKomp)
					{
					
						# Увеличиваем счетчик повторяющихся компьютеров.
						$KompRepeat ++
					}
					else
					{
						# Выполняем проверку на заполненость у сотрудника поля компания.
						if ($userlist[$NumberUser].Company -eq $null)
						{
							# Проверяем наличие имени офиса в массиве со всеми обработанами именами офиса.
							if($MassKOffice -eq $OfficeName)
							{
								$a = findinarr $MassKOffice $OfficeName
								$b = $MassKSum[$a]
								$b++
								$MassKSum[$a] = $b
							}
							else
							{
								$MassKOffice += $OfficeName
								$MassKSum += 1
							}
						}
						else
						{
							# Проверяем наличие имени офиса в массиве со всеми обработанами именами офиса.
							if($MassKompOffice -eq $OfficeName)
							{
								#Определяем массив со всеми индексами встречающегося в массиве элемента имени офиса.
								$Mas += findinarr $MassKompOffice $OfficeName
								
								# Проверяем наличие элемента с именем компании в массиве компаний по тем же индексам что и были найдены для массива офисов.
								# Нужно для того, что бы найти в 2 массивах с одинаковыми индаксами удовлетворяющие нас значения.
								if ($MassKompCompany[$Mas] -eq $userlist[$NumberUser].Company)
								{
									# Запишем значение в пустой массив, на тот случай если масив по индексу содержит всего один элемент.
									$Mass += $MassKompCompany[$Mas]
									# Записываем в переменную какой индекс имеет удовлетворяющее нас значени.
									$b = findinarr $Mass $userlist[$NumberUser].Company
									# Записываем в переменную индекс по которому в обоих масивах находятся удовлетворяющие нас условия. 
									$c = $Mas[$b]
									# Увеличиваем счетчик по такому же индексу.
									$b = $MassKompSum[$c]
									$b++
									$MassKompSum[$c] = $b
								}
								else
								{
									# Дописываем в конец массивов все интересующие нас значения
									$MassKompOffice += $OfficeName
									$MassKompCompany += $userlist[$NumberUser].Company
									$MassKompSum += 1
								}
							}
							else
							{
								# Дописываем в конец массивов все интересующие нас значения
								$MassKompOffice += $OfficeName
								$MassKompCompany += $userlist[$NumberUser].Company
								$MassKompSum += 1
							}
						}
					}
					# Добовляем имя компьютера в массив с обработанами именами.
					$MassKompName += $NameKomp
					
				}
				else						########################################################################################
				{
																			# ПОДСЧЕТ ЮЗЕРОВ
					
					# Если в заполняемом масиве с полем неопознаных подсетей для подсчета юзеров уже записано найденое неопознаная подсеть.
					if ($MassUserIP -eq $ip)
					{
						# Получаем индекс этого элемента офис и увеличиваем значение счетчика в масиве с подсчетом.
						$a = findinarr $MassUserIP $ip
						$b = $MassUserSumIP[$a]
						$b++
						$MassUserSumIP[$a] = $b
					}
					else
					{
						# Новые значения для счетчика в массив.
						$MassUserIP += $ip
						$MassUserSumIP += 1
					}
					
																			# ПОДСЧЕТ КОМПОВ
					
					# Проверяем наличие имени компьютера в массиве со всеми обработанами именами компьютеров.
					if($MassKompName -eq $NameKomp)
					{
						# Увеличиваем счетчик повторяющихся компьютеров.
						$KompRepeat ++
					}
					else
					{						
						# Увеличиваем счетчик компьютеров с неопознаной подсетью.
						$KompUnknown++
					}
					$MassKompName += $NameKomp
				}	
			}
		}
	}
	$ip = ""
	$NameKomp = ""
	$Mass = @()
	$Mas = @()
}

																		#Записываем все масивы в файл Excel.
																		

$p = 2

# Записываем лист "Подсчет" Имена офисов и счетчик.
For ($i = 0; $i -le $MassUserOffice.Count; $i++)
{
	$objWorkbook.Worksheets.Item($SheetName2).Cells.Item($p, 1).Value() = $MassUserOffice[$i]
	$objWorkbook.Worksheets.Item($SheetName2).Cells.Item($p, 2).Value() = $MassUserSumOffice[$i]
	$p++
}
$p = 2

# Записываем лист "НеопПодсети" Неопознаные подсети и счетчик.
For ($i = 0; $i -le $MassUserIP.Count; $i++)
{
	$objWorkbook.Worksheets.Item($SheetName3).Cells.Item($p, 1).Value() = $MassUserIP[$i]
	$objWorkbook.Worksheets.Item($SheetName3).Cells.Item($p, 2).Value() = $MassUserSumIP[$i]
	$p++
}
$p = 2

# Записываем лист "КолКомпов" Имя офиса, имя организации и счетчик.
For ($i = 0; $i -le $MassKompOffice.Count; $i++)
{
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($p, 1).Value() = $MassKompOffice[$i]
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($p, 2).Value() = $MassKompCompany[$i]
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($p, 3).Value() = $MassKompSum[$i]
	$p++
}

For ($i = 0; $i -le $MassKOffice.Count; $i++)
{
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($p, 1).Value() = $MassKOffice[$i]
	$objWorkbook.Worksheets.Item($SheetName4).Cells.Item($p, 3).Value() = $MassKSum[$i]
	$p++
}

$p = 2

# Записываем в Excel значения счетчиков компов с неопознаным офисом и компы с несколькими пользователями.
$objWorkbook.Worksheets.Item($SheetName4).Cells.Item(2, 5).Value() = $KompUnknown
$objWorkbook.Worksheets.Item($SheetName4).Cells.Item(2, 6).Value() = $KompRepeat


#Закрываем книгу Excel
$objExcel.Workbooks.close()
#Выходим из Excel (вернее даем команду на выход из Excel)
$objExcel.Quit()
#обнуляем объект
$objExcel = $null

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()