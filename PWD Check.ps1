#Program File Path
$DirPath = "C:\Users\dheydon\Documents\Projects\PasswordExpiry"
 
$Date = Get-Date
#Check if program dir is present
$DirPathCheck = Test-Path -Path $DirPath

Import-Module ActiveDirectory


#Number of days to notify user, if their password expires in x or less days
$expireindays = 14

# Get Users From AD who are Enabled, Passwords Expire, Part of company 'GS AUSTRALIA' and are Not Currently Expired
"$Date - INFO: Importing AD Module" | Out-File ($DirPath + "\" + "Log.txt") -Append
Import-Module ActiveDirectory
"$Date - INFO: Getting users" | Out-File ($DirPath + "\" + "Log.txt") -Append
$users = Get-Aduser -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress -filter { (Enabled -eq 'True') -and (PasswordNeverExpires -eq 'False') -and (company -eq 'GS AUSTRALIA') } | Where-Object { $_.PasswordExpired -eq $False }
 

foreach ($user in $users){
$Name = (Get-ADUser $user | ForEach-Object { $_.Name })
$firstName = (Get-ADUser $user | ForEach-Object { $_.givenName })
#Work out expiry
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
$passwordSetDate = (Get-ADUser $user -properties * | ForEach-Object { $_.PasswordLastSet })
$expireson = $passwordsetdate + $maxPasswordAge
$today = (get-date)

 #Store days to expire
 Try{ $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days }

  Catch { 
          Write-Warning "Failed to determine password expiry status for user $Name"
          "$Date - INFO: Failed to determine password expiry status for user $Name" | Out-File ($DirPath + "\" + "Log.txt") -Append
          Continue
        }


#Check days to expire < defined period of time to notify users (14 days)
     If (($daystoexpire -ge "0") -and ($daystoexpire -lt $expireindays)){
        $email = $user.EmailAddress
         "$daystoexpire, $email "|Out-File ($DirPath + "\" + "Log.txt") -Append

            #Email User
            $From = 'gsauhelpdesk@groupeseb.com'
            $Date = Get-Date -Format “dd.MM.yyyy”
            $Subject = 'Your password is expiring in ' + $daystoexpire + ' days.'
            $recipients = 'dheydon@groupeseb.com' #,'it.au@groupeseb.com'
            
			Send-MailMessage -BodyAsHtml -From $From -To $recipients -Subject $Subject -Body ($firstName + ' your password will expire in '+ ' '  + $daystoexpire +' day(s)'  + '. Please follow the below instructions relevant to you.<br /><br />' +'<b>Working From Home (Includes Users with both iPad and Laptop)</b><br /> Please login to the VPN on your laptop then follow the link below<br /> https://fs.seb.com/adfs/portal/updatepassword/ <br /><br /><b> In the Office </b><br />Please follow the link below to change your password. <br /> https://fs.seb.com/adfs/portal/updatepassword/<br />  <br / ><b> iPad Only </b><br />Please follow the link below to change your password. It may take up to 10 minutes for your password to update. <br /> https://fs.seb.com/adfs/portal/updatepassword/<br /> <br /> If you require assistance please email gsauhelpdesk@groupeseb.com <br / >') -SmtpServer 'smtp.seb.com'
            
			
			"$Date - INFO: Sent password expiry email for $Name" | Out-File ($DirPath + "\" + "Log.txt") -Append
            #Force user to change password if days to expire < 7
          #  $testing = $true
           #     if (!$testing){
             #       if($daystoexpire -lt 7){Set-Aduser $user -ChangePasswordAtLogon $true}
              #  }


     }
 }