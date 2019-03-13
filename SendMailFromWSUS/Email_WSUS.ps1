$smtp = "mail.craftmetall.ru"

$to = "shumov.du@craftmetall.ru"

$from = "Тех поддержка <support@group.com>"

$subject = "Профилактические работы на серверах компании"

$body = "<b><font color=red>Уважаемые Коллеги!</b></font> <br>"
$body += "Информируем Вас, что в ближайшую субботу в период времени с 18:00 до 22:00<br>"
$body += "на серверной инфраструктуре, будут производиться регламентные профилакимческие работы.<br>"
$body += "Во время работ, не будут доступны следующие сервисы:<br>"
$body += "1. Почтовый сервисы<br>"
$body += "2. Удаленный доступ<br>"
$body += "3. 1С<br><br><br><br><br>"
$body += "С уважением<br>"
$body += "Служба Технической поддержки<br>"
$body += "+7 499 455 03 17 / 1000 <br>"
$body += "support.craftmetall.ru / support@craftmetall.ru <br>"

#### Now send the email using \> Send-MailMessage 

send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Priority high

########### End of Script################
