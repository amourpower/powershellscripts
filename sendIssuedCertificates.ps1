<#
.SYNOPSIS
   Script is using certutil.exe to export certificates from CA and sends email if expiration date is lower than specified number of months.
.DESCRIPTION
   Script is using certutil.exe to export certificates from CA and sends email if expiration date is lower than specified number of months.
.PARAMETER -Months
   How many months before expiration date a notice should be sent.
.PARAMETER -NoMail
   Option to keep silent if errors occurred. No mail will be sent.
.EXAMPLE
   .\sendIssuedCertificates.ps1 -Months 1 -NoMail
   This Example will search for all issued certificates with expiration date lower than one month and shows results on the screen. Also will not send email if any matching certificates were found.
#>

# --------------------------------------------------
#Set all parameters
Param(
	[Int]$Months = $null,
    [switch]$noMail = $false,
    [string]$mailstatus = $null
)
# --------------------------------------------------
#functions

Function Send-CertificateList
{
    $FromAddress = 'eivastausta@aia.com'
    $ToAddress = 'aaron-zx.huang@aia.com'
    $MessageSubject = "Certificate expiration reminder from $env:COMPUTERNAME.$env:USERDNSDOMAIN"
    $SendingServer = 'smtphk-int.aia.biz'

    $SMTPMessage = New-Object System.Net.Mail.MailMessage $FromAddress, $ToAddress, $MessageSubject,$mailbody -ErrorAction SilentlyContinue
    $SMTPMessage.IsBodyHTML = $true
    $SMTPMessage.Priority = [System.Net.Mail.MailPriority]::High
    $SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer -ErrorAction SilentlyContinue

    if(Test-Connection -Cn $SendingServer -BufferSize 16 -Count 1 -ea 0 -quiet){
	    $SMTPClient.Send($SMTPMessage)
    }else{
	    Write-Host 'No connection to SMTP server. Failed to send email!'
        Write-Output 'No connection to SMTP server. Failed to send email!' | Out-File $mailstatus -Append
    }
}

# --------------------------------------------------

#HTML Style
$style = "<style>body{font-family:`"Calibri`",`"sans-serif`"; font-size: 14px;}"
$style = $style + "@font-face
	{font-family:`"Cambria Math`";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;}"
$style = $style + "table{border: 1px solid black; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;}"
$style = $style + "th{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "td{border: 1px solid black; padding: 5px; }"
$style = $style + ".crtsn{font-weight: bold; color: blue; }"
$style = $style + ".crtexp{font-weight: bold; color: red; }"
$style = $style + ".crtcn{font-weight: bold; color: orange; }"
$style = $style + "</style>"

# --------------------------------------------------
#variables
$strDate = Get-Date -format yyyyMMdd
$exportFileName = "certificates_" + $strDate + ".csv"
$now = Get-Date
$nowm = $now.Month
$nowy = $now.Year
$mailbody = @()
$expirymy = @()
$table = @()

# --------------------------------------------------
#variables

#export certificates to CSV
certutil.exe -view csv > $exportFileName

#Import certificates from CSV
$importexp = Import-Csv $exportFileName | Select-Object 'Certificate Expiration Date'
$importall = Import-Csv $exportFileName | Where-Object {$_.'Serial Number' -notcontains 'EMPTY'} | Select-Object -Property 'Request ID','Serial Number','Requester Name','Certificate Expiration Date','Request Common Name','Request Disposition'

#build email body in HTML
$mailbody += '<html><head><meta http-equiv=Content-Type content="text/html; charset=utf-8">' + $style + '</head><body>'
$mailbody += "Theese certificates will expire soon:<br />"

#convert dates and extract month and year
foreach($i in $importall){
   $expiry = Get-Date $i.'Certificate Expiration Date' | Select Month, Year
   $expirymy += $expiry
}

#cycle through array and search for matching cetificates
for($i=0;$i -lt $expirymy.Count;$i++){
    if(($expirymy[$i].Month -gt $nowm) -and ($importall[$i].'Request Disposition' -contains '20 -- issued')){
        if((($expirymy[$i].Month - $nowm) -le $Months) -and (($expirymy[$i].Year - $nowy) -eq 0)){
            Write-Host 'Certificate ID:' $importall[$i].'Request ID' 'with Serial Number:' $importall[$i].'Serial Number' 'will expire in ' -NoNewline; Write-Host ($expirymy[$i].Month - $nowm) 'months!'-ForegroundColor Red
            Write-Host 'This certificate has DN: ' -NoNewline; Write-Host $importall[$i].'Request Distinguished Name' -ForegroundColor DarkYellow
            Write-Host 'Please don`t forget to renew this certificate before expiration date: ' -NoNewline; Write-Host $importall[$i].'Certificate Expiration Date' -ForegroundColor Red "`n"
            
            $mailbody += '<p>'
            $mailbody += 'Certificate ID: ' + $importall[$i].'Request ID' + ' with Serial Number: <span class="crtsn"">' + $importall[$i].'Serial Number' + '</span> will expire in <span class="crtexp">' + ($expirymy[$i].Month - $nowm) + ' months!</span>'+"<br />"
            $mailbody += 'This certificate has CN: <span class="crtcn">' + $importall[$i].'Request Common Name' + "</span><br />"
            $mailbody += 'Please don`t forget to renew this certificate before expiration date: <span class="crtexp">' + $importall[$i].'Certificate Expiration Date'+"</span>"
            $mailbody += '</p>'
            $table += $importall[$i] | Sort-Object 'Certificate Expiration Date' | Select-Object -Property 'Request ID','Serial Number','Requester Name','Certificate Expiration Date','Request Common Name'
        }
    }
}

$mailbody += '<p><table>'
$mailbody += '<th>Request ID</th><th>Serial Number</th><th>Requester Name</th><th>Requested CN</th><th>Expiration date</th>'
foreach($row in $table){
    $mailbody += "<tr><td>" + $row.'Request ID' + "</td><td>" + $row.'Serial Number' + "</td><td>" + $row.'Requester Name' + "</td><td>" + $row.'Request Common Name' + "</td><td>" + $row.'Certificate Expiration Date' + "</td></tr>"
}
$mailbody += '</table></p>'
$mailbody += '</body>'
$mailbody += '</html>'

#if there are matching certificates found send email
if(($table.Count -gt '0') -and (!$noMail)){
    Send-CertificateList
}

#remove CSV file
Remove-Item $exportFileName