cls
#Убивает процесс если он открыт в противном случае выдает ошибку
(Get-WmiObject win32_process -Filter "name='cmd.exe'" -ComputerName "cp777").Terminate() 

#Запуск службы на удаленном компе
(Get-WmiObject win32_Service -Filter "name='WinRM'" -ComputerName "cp555").StartService()

# Перевод служб в другой режим запуска.
Set-Service -Name WinRM -ComputerName $Comp -StartupType Manual