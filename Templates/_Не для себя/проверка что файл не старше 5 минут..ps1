cls

									# Переменные
# Путь к файлу.
$Files = "E:\123\123.txt"
# Адрес кому отправляется почта.
$ToSend = ""
# Адрес кому отправляется почта.
$FromSend = ""
# Тема письма.
$SubjectSend = ""
# Сервер SMTP.
$SmtpServer = ""
# Содержание письма.
$BodySend = ""


									# Тело скрипта
									

# Получаем время изменения объекта.
$FilesCreationTime = Get-ItemProperty -Path $Files | ForEach-Object {$_.LastWriteTime} 

# Проверяем старше ли файл 5 минут.
if ($FilesCreationTime -lt (Get-Date).AddMinutes(-5))
{
 Send-MailMessage -From $FromSend -To $ToSend -Subject $SubjectSend -SmtpServer $SmtpServer -Body $BodySend
}