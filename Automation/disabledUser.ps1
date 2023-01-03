Import-Module ActiveDirectory

# Get all disabled users
$disabledUsers = Get-ADUser -Filter {Enabled -eq $false -and GivenName -eq 'Jacob'} -SearchBase "OU=AU-SYD-Disabled Objects-Users,OU=AU-SYD-Disabled Objects,OU=AU-SYD,OU=AU, DC=as,DC=seb,DC=com"

# Iterate through each disabled user
foreach ($user in $disabledUsers) {
  # Update the email auto-reply for the user
  Set-MailboxAutoReplyConfiguration -Identity $user.SamAccountName -AutoReplyState Enabled -InternalMessage "We are sorry but this email account you have tried to use does not exist within our organisation." -ExternalMessage "We are sorry but this email account you have tried to use does not exist within our organisation."
}

