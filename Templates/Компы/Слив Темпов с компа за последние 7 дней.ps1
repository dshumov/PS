cls

$comp = "cp021"
$user = ""
$FilesName = ""

$user = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName "$comp").UserName
$user = $user -split "\", 0, "simplematch"
$user = $user[1..1]

New-Item E:\Слив\$user\Temp -ItemType Directory

Get-ChildItem \\$comp\c$\Users\$user\appdata\local\temp -Recurse | 
Where-Object {$_.LastAccessTime -lt (Get-Date).AddMonths(-1)} |
ForEach-Object {Remove-Item -Path $_.FullName -Recurse} 

#-and $_.FullName -inotlike "**"

Get-ChildItem \\$comp\c$\Users\$user\appdata\local\temp -Filter "*.bak" -Recurse | Where-Object {$_.FullName -inotlike "*Автобат*"`
-and $_.FullName -inotlike "*Балашиха*" -and $_.FullName -inotlike "*Бескудниково*" -and $_.FullName -inotlike "*Windows*"`
-and $_.FullName -inotlike "*Бобров*" -and $_.FullName -inotlike "*Бутово*" -and $_.FullName -inotlike "*Видное*" `
-and $_.FullName -inotlike "*Дрожжино*" -and $_.FullName -inotlike "*Дрожино*" -and $_.FullName -inotlike "*Катюшки*" `
-and $_.FullName -inotlike "*Красная*" -and $_.FullName -inotlike "*Лечебная*" -and $_.FullName -inotlike "*Новский*" `
-and $_.FullName -inotlike "*Путилково*" -and $_.FullName -inotlike "*Сакраменто*" -and $_.FullName -inotlike "*Северное*" `
-and $_.FullName -inotlike "*Северный*" -and $_.FullName -inotlike "*Середнево*" -and $_.FullName -inotlike "*Солнцево*" `
-and $_.FullName -inotlike "*Щитниково*" -and $_.FullName -inotlike "*Южное*" -and $_.FullName -inotlike "*Кучино*" `
-and $_.FullName -inotlike "*Пыхтино*" -and $_.FullName -inotlike "*Учеба*" -and $_.FullName -inotlike "*Каталог*" `
-and $_.FullName -inotlike "*Ольгино*" -and $_.FullName -inotlike "*Сокраменто*" -and $_.FullName -inotlike "*Program Files*" `
-and $_.FullName -inotlike "*.tmp*" -and $_.LastAccessTime -gt (Get-Date).AddDays(-7)} |

ForEach-Object {$FullFilesName = $_.Name; `
Copy-Item $_.FullName E:\Слив\$user\Temp -Recurse; `
$FilesName = $_.Name -replace "_1_\S+"; `
$FilesName = "$FilesName" + ".dwg"; `
Rename-Item -Path E:\Слив\$user\Temp\$FullFilesName -NewName $FilesName}

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()