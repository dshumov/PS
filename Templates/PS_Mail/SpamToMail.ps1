$date_x=(get-date -uformat "%u")
$date_n=(get-date((get-date).adddays(-$date_x))).tostring("dd'/'MM'/'yyyy");
$date_y=(get-date((get-date).adddays(-6).adddays(-$date_x))).tostring("dd'/'MM'/'yyyy");
$rep=(Get-AgentLog -EndDate $date_n -StartDate $date_y | group action | sort-object Name | ft name,count -Autosize);
$SmtpClient = new-object system.net.mail.smtpClient;
$MailMessage = New-Object system.net.mail.mailmessage;
$SmtpClient.Host = "mail.vth-invest.ru";
$mailmessage.from = "postmaster@vth-invest.ru";
$mailmessage.To.add("s.morozov@vth-invest.ru");
$mailmessage.Subject = "Отчет по спаму c " + $date_y + " по " + $date_n ;
$msg = $rep | format-table | out-string -width 160;
$mailmessage.Body = $msg;
$smtpclient.Send($mailmessage);