$maxdays=(Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.TotalDays
$summarybody="Name `t `t `t ExpireDate `t `t `t DaysToExpire `n"
 
(Get-ADUser -filter {(Enabled -eq "True") -and (PasswordNeverExpires -eq "False") -and (emailaddress -like "*") -and (objectclass -eq "user")} -properties *) | Sort-Object pwdLastSet |
foreach-object {
 
$lastset=Get-Date([System.DateTime]::FromFileTimeUtc($_.pwdLastSet))
$expires=$lastset.AddDays($maxdays).ToShortDateString()
$daystoexpire=[math]::round((New-TimeSpan -Start $(Get-Date) -End $expires).TotalDays)
$samname=$_.samaccountname
$firstname=$_.GivenName
$emailaddress=$_.emailaddress
if ($daystoexpire -le 9){
    $ThereAreExpiring=$true
 
    $emailFrom = "@hotmail.com"
    $emailTo = "$emailaddress"
    $subject = "$firstname, your password expires in $daystoexpire day(s)"
    $body = "$samname, $emailaddress
    Your password expires in $daystoexpire day(s).
    Please change your password as soon as convenient so your access will not be denied. Also please remember to change your password for any accounts synced on your iOS and Android devices."
    $smtpServer = "mx..com"
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($emailFrom, $emailTo, $subject, $body)
 
    $summarybody += "$samname `t `t `t $expires `t `t $daystoexpire `n"
}
}
if ($ThereAreExpiring) {
$emailFrom = "IT@.com"
$emailTo = "support@.com"
$subject = "Expiring passwords"
$body = $summarybody
$smtpServer = "mx.e.com"
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($emailFrom, $emailTo, $subject, $body)
}
