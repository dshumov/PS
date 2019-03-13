cls

											# ПЕРЕМЕННЫЕ
# Вспомогательные переменные
$ReceiveJob = ""
$OutData = ""
$HDD = ""
# Переменные масивов получаемые скриптом с ПК
$OutMassComps = @()
$OutMassUsers = @()
$OutMassProccessor = @()
$OutMassMemory = @()
$OutMassHDD = @()
$OutMassLogicalD = @()
$OutMassVideo = @()
# Переменные масивов получаемые из Экселя
$InMassComps = @()
$InMassUsers = @()
$InMassProccessor = @()
$InMassMemory = @()
$InMassHDD = @()
$InMassVideo = @()
$InMassDate = @()
# Переменные масивов для новых пользователей и их ПК
$NewMassComps = @()
$NewMassUsers = @()
$NewMassProccessor = @()
$NewMassMemory = @()
$NewMassHDD = @()
$NewMassLogicalD = @()
$NewMassVideo = @()
# Вспомогательные переменные масивов.
$PastMass = @()
$BeginMass = @()
$IndexUser = @()
$IndexUpdate = @()
$UsersAD = @()
$Users = @()
$IndexUsersNull = @()

get-Job | Remove-Job -Force


# Получаем полный список компьютеров из каталога СР14
$UsersAD = Get-QADUser -SearchRoot "morton.ru/_Мортон/Учётные записи/СР14/" -Enabled

# Получаем полный список компьютеров из каталога СР14
$Complist =  Get-QADComputer -SearchRoot "morton.ru/_Мортон/Ресурсы/Рабочие станции/СР14/" -IncludedProperties "name"
# Перебор компов в массиве.
for ($i=0; $i -ne $Complist.Count; $i++)
{
	# переменная для выхода из безконечного цикла
	$Check = $false
	while ($Check -eq $false) 
	{
		if ((Get-Job -State 'Running').Count -lt 15)
		{
	
			# Присваиваем переменной значения ячейки массива
			$Comp = $Complist[$i].name
			$Comp
			

			Start-Job -InitializationScript {Add-pssnapin Quest.ActiveRoles.ADManagement} -ArgumentList $Comp -ScriptBlock {
					
				$Comp = $args
				
				# Проверка активности компьютера.
				$c = Test-Connection -ComputerName $Comp -Count 1 -Quiet
				
				# Если Комп включен, выполняем копирование и запуск файла. 
				if ($c -eq "True")
				{
					
					# Добовляем компьютер в массив с именами компов
					$OutComps = $Comp
					

					# Текущий пользователь 
					$user = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName $comp).UserName
					if ($user -ne $null)
					{
						# Получение полного ФИО пользователя
						$OutUsers = (Get-QADUser -SearchRoot "morton.ru/_Мортон/Учётные записи/СР14/" -Identity $user).Name
						

						# Выборка процесоров
						$OutProccessor = ""
						Get-WmiObject -Class Win32_Processor -ComputerName $Comp | ForEach-Object {$OutProccessor = $OutProccessor + $_.Name + "`n"}
						#$OutProccessor
						

						# Выборка ОЗУ подсчет количества планок и общий объем
						$FullMemory = 0
						$AmountMemory = 0
						Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $Comp| 
						ForEach-Object {$FullMemory = $FullMemory + ($_.Capacity -split "", 0, "simplematch")[1]; $AmountMemory = $AmountMemory + 1}
						$OutMemory = "$AmountMemory" + " Шт. `nВсего " + "$FullMemory" + "Gb"
						#$OutMemory


						# Выбока Физических дисков
						$OutHDD = ""
						Get-WmiObject -Class Win32_DiskDrive -ComputerName $Comp -Filter "MediaType = 'Fixed hard disk media'" |
						ForEach-Object {$OutHDD += "Disk" + $_.Index + " = " + ($_.size / 1000000000 -split '[.]')[0] + "`n"}
						
						# Выбока логических дисков
						$OutLogicalD = ""
						Get-WmiObject -Class Win32_LogicalDisk -ComputerName $Comp -Filter "Description = 'Local Fixed Disk'" |
						ForEach-Object {$OutLogicalD += $_.Name +  " = " + ($_.size / 1000000000 -split '[.]')[0] + "`n"}


						# Выборка Видеокарточки
						$OutVideo = ""
						Get-WmiObject -Class Win32_VideoController -ComputerName $Comp | Where-Object {$_.Name -inotlike "DameWare*"} |
						ForEach-Object {$OutVideo = $OutVideo + ($_.Name -split "(", 0, "simplematch")[0] + "`n"}
						#$OutVideo
						
						#Собираем все полученые данные в переменую
						$OutData = $OutComps + $OutUsers + $OutProccessor + $OutMemory + $OutHDD + $OutLogicalD + $OutVideo
						$OutData
					}
				}
			}
			$Check = $true
		}
	}

}

# Ждем пока не закончица выполнение всех потоков 
while ((get-Job) -ne $null)
{
	# Передаем результаты всех скриптов в массивы.
	get-Job -State 'Completed' | 
			foreach {$ReceiveJob = Receive-Job -Name $_.Name
			if ($ReceiveJob -ne $null)
			{
			$OutMassComps += ($ReceiveJob -split ",")[0]
			$OutMassUsers += $ReceiveJob[1]
			$OutMassProccessor += $ReceiveJob[2]
			$OutMassMemory += $ReceiveJob[3]
			$OutMassHDD += $ReceiveJob[4]
			$OutMassLogicalD += $ReceiveJob[5]
			$OutMassVideo += $ReceiveJob[6]
			}
			Remove-Job -Name $_.Name
		}
}

#задаем путь к файлу Excel
$FileExcel = "J:\IT\Учёт\Склад\Инвентаризация компов PS.xlsx"
$ExcelListKomps = "Компы"
$ExcelListUpdate = "Update"
$ExcelListZamena = "Выдача и замена ПК"


#открываем Excel
$ExcelObject = New-Object -ComObject Excel.Application
$ExcelObject.Visible = $true
$WorkbookExcel = $ExcelObject.Workbooks.Open($FileExcel)
$SheetExcel = $WorkbookExcel.worksheets.Item($ExcelListKomps)

# Получаем последнюю используемую строку
$TotalRows = $WorkbookExcel.worksheets.Item($ExcelListKomps).UsedRange.Rows.count

#Выгрузка данных из экселя кроме Логических разделов.
$WorkbookExcel.worksheets.Item($ExcelListKomps).Range("A2:A$TotalRows").value() | ForEach-Object {$InMassComps += @($_)} | Out-Null
$WorkbookExcel.worksheets.Item($ExcelListKomps).Range("B2:B$TotalRows").value() | ForEach-Object {$InMassProccessor += @($_)} | Out-Null
$WorkbookExcel.worksheets.Item($ExcelListKomps).Range("C2:A$TotalRows").value() | ForEach-Object {$InMassMemory += @($_)} | Out-Null
$WorkbookExcel.worksheets.Item($ExcelListKomps).Range("D2:B$TotalRows").value() | ForEach-Object {$InMassHDD += @($_)} | Out-Null
$WorkbookExcel.worksheets.Item($ExcelListKomps).Range("E2:A$TotalRows").value() | ForEach-Object {$InMassVideo += @($_)} | Out-Null
$WorkbookExcel.worksheets.Item($ExcelListKomps).Range("G2:A$TotalRows").value() | ForEach-Object {$InMassUsers += @($_)}  | Out-Null
$WorkbookExcel.worksheets.Item($ExcelListKomps).Range("I2:A$TotalRows").value() | ForEach-Object {$InMassDate += @($_)} | Out-Null


#_________________________________________________________________#############################_____________________________________###


# Перебираем компьюторы полученые скриптом
for ($i=0; $i -lt $OutMassComps.Count; $i++)
{
	# Если комп есть в массиве выгруженом из Экселя
	if ($InMassComps -contains $OutMassComps[$i])
	{	
		# Перебираем компьютеры полученые из Экселя
		for ($n=0; $n -lt $InMassComps.Count; $n++)
		{
			# Если мы нашли нужный компьютер
			if ($InMassComps[$n] -eq $OutMassComps[$i])
			{
				# Если пользователь полученый скриптом не равен пользователю из экселя
				if ($InMassUsers[$n] -ne $OutMassUsers[$i])
				{
					# Записываем данные индексов нахождения ПК в массивах.
					$IndexUser += "$i,$n"   
				}
				
				# Сравниваем начинку компов для того что бы убедится что не чего не изменилось.
				# Для общего контроля изминений заводим переменую и присваиваем ей значение $False
				# В случае если начинка ПК была изменина переменной присвоится значение $true для последующей записи всех изминений.
				
				$Past = ""
				$Begin = ""
				
				$Check = $false
				
				# Сравнение процессоров
				if ($OutMassProccessor[$i] -ne $InMassProccessor[$n])
				{
					# Записываем старые и новые параметры.
					$Past += $InMassProccessor[$n] + "`n" + "`n"
					$Begin += $OutMassProccessor[$i] + "`n" + "`n"
					# Меняем переменую для дальнейшего подытоживания всех изминений.
					$Check = $true
				}
				
				# Сравнение памяти
				if ($OutMassMemory[$i] -ne $InMassMemory[$n])
				{
					# Записываем старые и новые параметры.
					$Past = $InMassMemory[$n] + "`n" + "`n"
					$Begin = $OutMassMemory[$i] + "`n" + "`n"
					# Меняем переменую для дальнейшего подытоживания всех изминений.
					$Check = $true
				}
				
				# Сравнение дисков
				if ($OutMassHDD[$i] -ne $InMassHDD[$n])
				{
					# Записываем старые и новые параметры.
					$Past = $InMassHDD[$n] + "`n" + "`n"
					$Begin = $OutMassHDD[$i] + "`n" + "`n"
					# Меняем переменую для дальнейшего подытоживания всех изминений.
					$Check = $true
				}
				
				# Сравнение видео
				if ($OutMassVideo[$i] -ne $InMassVideo[$n])
				{
					# Записываем старые и новые параметры.
					$Past = $InMassVideo[$n]
					$Begin = $OutMassVideo[$i]
					# Меняем переменую для дальнейшего подытоживания всех изминений.
					$Check = $true
				}
				if ($Check -eq $true)
				{
					# Записываем все результаты в конечные массивы
					$IndexUpdate += "$i,$n"
					$PastMass += $Past
					$BeginMass += $Begin
				}
			}
		}
	}
	else
	{
		# Так как комп новый описываем установку ПК
		$NewMassComps += $OutMassComps[$i]
		$NewMassProccessor += $OutMassProccessor[$i]
		$NewMassMemory += $OutMassMemory[$i]
		$NewMassHDD += $OutMassHDD[$i]
		$NewMassLogicalD += $OutMassLogicalD[$i]
		$NewMassVideo += $OutMassVideo[$i]
		$NewMassUsers += $OutMassUsers[$i]
		
	}
}

# Проверим что пользователи из экселя есть в выгруженых пользователях из АД
$UsersAD | ForEach-Object {$Users += $_.Name}

#Перебираем юзеров выгруженных из Экселя
for ($i=0; $i -lt $InMassUsers.Count; $i++)
{
	# Если Юзера нету в масиве полученом из АД
	if($Users -notcontains $InMassUsers[$i])
	{
		$IndexUsersNull += $i
	}
}


# Получаем последнюю используемую строку на листе
$TotalRows = $WorkbookExcel.worksheets.Item($ExcelListKomps).UsedRange.Rows.count



#________________________________________________ Записываем в Эксель новые ПК с Юзером
												
For ($i = 0; $i -lt $NewMassComps.Count; $i++)
{
	$TotalRows++
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 1).Value() = $NewMassComps[$i]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 2).Value() = $NewMassProccessor[$i]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 3).Value() = $NewMassMemory[$i]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 4).Value() = $NewMassHDD[$i]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 5).Value() = $NewMassLogicalD[$i]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 6).Value() = $NewMassVideo[$i]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 7).Value() = $NewMassUsers[$i]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($TotalRows, 9).Value() = (Get-Date -DisplayHint Date)
	
	# Запись в лист "Выдача и замена ПК" будет ниже в соответствующем блоке скрипта.
}


#________________________________________________ Записываем в Эксель информацию об Update компьютера.

# Получаем последнюю используемую строку на листе
$TotalRows = $WorkbookExcel.worksheets.Item($ExcelListUpdate).UsedRange.Rows.count												

For ($i = 0; $i -lt $IndexUpdate.Count; $i++)
{
	# Разбиваем элемент масива на индексы.
	$Indexs = $IndexUpdate[$i] -split ","
	# Заполняем лист Update
	$TotalRows++
	$WorkbookExcel.worksheets.Item($ExcelListUpdate).Cells.Item($TotalRows, 1).Value() = $OutMassComps[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListUpdate).Cells.Item($TotalRows, 2).Value() = $PastMass[$i]
	$WorkbookExcel.worksheets.Item($ExcelListUpdate).Cells.Item($TotalRows, 3).Value() = $BeginMass[$i]
	$WorkbookExcel.worksheets.Item($ExcelListUpdate).Cells.Item($TotalRows, 4).Value() = (Get-Date -DisplayHint Date)
	
	# Вносим изминения в лист Компы.
	# В переменую $n записываем строку где записаны старые параметры компа в эксель.
	$n = $InMassComps[$Indexs[1]] + 2
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 1).Value() = $OutMassCompsм[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 2).Value() = $OutMassProccessor[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 3).Value() = $OutMassMemory[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 4).Value() = $OutMassHDD[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 5).Value() = $OutMassLogicalD[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 6).Value() = $OutMassVideo[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 7).Value() = $OutMassUsers[$Indexs[0]]
	$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 8).Value() = (Get-Date -DisplayHint Date)
}


#________________________________________________ Записываем в Эксель информацию о выдаче и замене ПК.

# Получаем последнюю используемую строку на листе
$TotalRows = $WorkbookExcel.worksheets.Item($ExcelListZamena).UsedRange.Rows.count												

# Проверяем что есть новые компы
if($NewMassComps -ne $null)
{
	$TotalRows++
	$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 1).Value() = $NewMassComps[$i]
	$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 2).Value() = $NewMassProccessor[$i]
	$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 3).Value() = $NewMassMemory[$i]
	$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 4).Value() = $NewMassHDD[$i]
	$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 5).Value() = $NewMassVideo[$i]
	$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 6).Value() = $NewMassUsers[$i]
	$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 7).Value() = (Get-Date -DisplayHint Date)
}

# Проверяем что есть компы с новыми пользователями.
if($IndexUser -ne $null)
{
	# Перебором выбираем все индексы из массивов.
	For ($i = 0; $i -lt $IndexUser.Count; $i++)
	{		
		$Indexs = $IndexUser[$i] -split ","
		
		if($OutMassComps -notcontains $InMassComps[$Indexs[1]])
		{
			#Объеденяем ячейки для определения компьютера на склад.
			$TotalRows++
			$n = $TotalRows++
			
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("A$TotalRows:A$n").Merge()
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("B$TotalRows:B$n").Merge()
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("C$TotalRows:C$n").Merge()
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("D$TotalRows:D$n").Merge()
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("E$TotalRows:E$n").Merge()
			
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 1).Value() = $InMassComps[$Indexs[1]]
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 2).Value() = $InMassProccessor[$Indexs[1]]
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 3).Value() = $InMassMemory[$Indexs[1]]
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 4).Value() = $InMassHDD[$Indexs[1]]
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 5).Value() = $InMassVideo[$Indexs[1]]
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 6).Value() = $InMassUsers[$Indexs[1]]
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 7).Value() = $InMassDate[$Indexs[1]]
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($n, 6).Value() = "Склад"
			$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($n, 7).Value() = (Get-Date -DisplayHint Date)
			
		}
		
		$TotalRows++
		$n = $TotalRows++
		
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("A$TotalRows:A$n").Merge()
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("B$TotalRows:B$n").Merge()
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("C$TotalRows:C$n").Merge()
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("D$TotalRows:D$n").Merge()
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Range("E$TotalRows:E$n").Merge()
		
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 1).Value() = $OutMassComps[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 2).Value() = $OutMassProccessor[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 3).Value() = $OutMassMemory[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 4).Value() = $OutMassHDD[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 5).Value() = $OutMassVideo[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 6).Value() = $InMassUsers[$Indexs[1]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($TotalRows, 7).Value() = $InMassDate[$Indexs[1]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($n, 6).Value() = $OutMassUsers[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListZamena).Cells.Item($n, 7).Value() = (Get-Date -DisplayHint Date)
		
		# Вносим изминения в лист Компы.
		# В переменую $n записываем строку где записаны старые параметры компа в эксель.
		$n = $InMassComps[$Indexs[1]] + 2
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 1).Value() = $OutMassCompsм[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 2).Value() = $OutMassProccessor[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 3).Value() = $OutMassMemory[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 4).Value() = $OutMassHDD[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 5).Value() = $OutMassLogicalD[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 6).Value() = $OutMassVideo[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 7).Value() = $OutMassUsers[$Indexs[0]]
		$WorkbookExcel.worksheets.Item($ExcelListKomps).Cells.Item($n, 8).Value() = (Get-Date -DisplayHint Date)
	}
}

#Сохраняем книгу Excel
$WorkbookExcel.Save()
#Закрываем книгу Excel
$WorkbookExcel.close()
#Выходим из Excel (вернее даем команду на выход из Excel)
$ExcelObject.Quit()
#обнуляем объект
$ExcelObject = $null

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()