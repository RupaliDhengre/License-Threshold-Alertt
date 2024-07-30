#Connect to the MSOnline Module
Connect-MsolService
#Give the Path name as per your requirement
$path= "C:\Users\Desktop"
#Read-Host -Prompt “Enter your tenant Global Admin password” -AsSecureString | ConvertFrom-SecureString | Out-File "$path\TENANTNAME.key"
$username = "abc@domainName.com"
$password = cat "$path\credential.key" | ConvertTo-SecureString
$TenantCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
Connect-MsolService -Credential $TenantCredentials

$smtpsettings = @{
To = "Enter recipient email address"
From = "Enter Sender Email Address"
#Enter the SMTP Server details in below section.
    
SmtpServer = "smtp.office365.com"
    Port = "587"
    Subject = "Alert Notification License Usage"
}

#Below variable mentions the license alert limit, if the assigned license counts reaches 90% of available licenses, it will generate the alert notification.
$key = 0.90

$array = @()
$row=@()
$newline = "<br>"
$htmlhead = "<html>
    <style>
BODY{font-family: Arial; font-size: 8pt;}
H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
TABLE{border: 2px solid black; border-collapse: collapse; font-size: 8pt;}
TH{border: 2px solid #000000; background: #cccccc; padding: 12px; color: #000000;}
TD{border: 2px solid #000000; padding: 5px; }
td.pass{background: #B7EB83;}
td.warn{background: #FFF275;}
td.fail{background: #FF2626; color: #ffffff;}
td.info{background: #85D4FF;}
</style>
<body>
    <p> Hello Team,$newline $newline
A license (or more) has a usage rate above the defined limit of $($key*100)%$newline</p>
"

$htmltail = "</body></html>"
#endregion Variable

#----------------------------------------------
#region Main Script
#----------------------------------------------

#Connect-MsolService -Credential $cred
#Set the license variable
$Licenses = Get-MsolAccountSku

Foreach ($License in $Licenses) {
    $FreeLicense = $License.ActiveUnits - $License.ConsumedUnits
    $Percentageusage = [math]::Round(($License.ConsumedUnits / $License.ActiveUnits)*100)


##Below code snippet will check if the License number has reached the threshold.

    if (($License.ConsumedUnits / $License.ActiveUnits) -ge $key) {
        if ($LicensesNotMonitor -notcontains $License.AccountSKuId){
            if (($LicensesSpecifics -notcontains $License.AccountSKuId) -OR ($FreeLicense -lt 1000)) {
                #write-output "Alert for license $($License.AccountSkuId) the usage rate of this license is greater than $($key*100)% it is currently $Percentageusage %"
                $row = New-object PSobject -Property @{
                   
                   
                     AvailableLicense = $FreeLicense
                    LicenseName = $License.SkuPartNumber
                    TotalLicense = $License.ActiveUnits
                   PercentageUsage = [math]::Round($Percentageusage)
                   
                }
                $array += $row
            }
        }
    }
}

if ($array.count -gt 0) {
    $html = $array | ConvertTo-Html -Property LicenseName,TotalLicense,AvailableLicense,PercentageUsage
    $body = $htmlhead + $html + $htmltail
    Send-MailMessage @smtpsettings -body $body -BodyAsHtml -Credential (Get-Credential)

}
