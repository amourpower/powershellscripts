$html = get-Content '.\emailbody.html'

foreach ($i in $html){ $body += $i }

Send-MailMessage -Credential amourpower@163.com -From amourpower@163.com -To amourpower@163.com -Subject FirstMail -SmtpServer smtp.163.com -Body $body -BodyAsHtml