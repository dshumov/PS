cls

$user = ""
$FilesName = ""

$user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
$user = $user -split "\", 0, "simplematch"
$user = $user[1..1]

Get-ChildItem C:\Users\$user\appdata\local\temp -Recurse | 
Where-Object {$_.LastAccessTime -lt (Get-Date).AddDays(-7)} |
ForEach-Object {Remove-Item -Path $_.FullName -Recurse -Force} 

#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()