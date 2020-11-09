# Stephanie Seyler
# 2020-10-27
# v1.0.1

# Build runtime location
$homeFolder = (split-path -path $PSScriptRoot)
Set-Location -Path $homeFolder

# Build Log object and headers
try {
    $errorcount = 0
    $LogLocation = $homefolder +  "\logs\" + (get-date -UFormat "%Y-%m-%d") + " Safe Sender log.csv"
    $log = new-object System.IO.StreamWriter ($Loglocation,$false,(new-object System.Text.UTF8Encoding($true)))
    $log.WriteLine('"Date/Time","Type","Information"')
    $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""BEGIN"",""Begin Logging""")
}catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to create new log file""");$errorcount++}

# Build M365 Credential object, export to Credentials folder
try {
    $CredentialLocation = $homeFolder + "\credentials\m365credentials.xml"
    # Test if credentials exist, if not prompt to create them
    if((test-path $CredentialLocation) -eq $false){
        $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""INFO"",""Missing credential: prompted to create""")
        get-credential | Export-Clixml $CredentialLocation
    }
    $UserCredential = Import-Clixml $CredentialLocation
    $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Sucess"",""Imported Credential: $($usercredential.username)""")
}catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to retrieve User Credential for Exchange online""");$errorcount++}

# Build list of users to update from data.csv and Active Directory
try {
    # Import CSV with group names and actions to be taken
    $data = Import-Csv ($homeFolder + "\data\data.csv")
    # Build DataTable object to hold Userprincipalname and what additions to be made to lists
    $Users = New-Object System.Data.Datatable; [void]$Users.Columns.Add("Userprincipalname")
    [void]$Users.Columns.Add("AddSafeSender"); [void]$Users.Columns.Add("AddBlockedSender")
    [void]$Users.Columns.Add("RemoveSafeSender"); [void]$Users.Columns.Add("RemoveBlockedSender")
    foreach($group in $data){
        $GroupUsers = Get-ADGroupMember -identity $group.groupname
        foreach($name in $GroupUsers){
            $account = Get-ADUser -Identity $name.SamAccountName
            # Add userprincipalname and the Fields from the associated group
            [void]$Users.rows.Add($account.userprincipalname, $group.AddSafeSender, $group.AddBlockedSender, `
            $group.RemoveSafeSender, $group.RemoveBlockedSender)
        }
    }
}catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to Build Users list""");$errorcount++}

# Connecting to exchange online using $userCredential imported from $CredentialLocation
try {
    Connect-ExchangeOnline -Credential $UserCredential -ShowProgress $true
    $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Sucess"",""Connected to Exchange Online""")
}catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to Connect to Exchange Online""");$errorcount++}

# TODO Build the Safe senders list
try {
    foreach($user in $users){
        # AddSafeSender
        try {
            if($user.AddSafeSender -notlike "Null"){
                Set-MailboxJunkEmailConfiguration $user.userprincipalname -TrustedSendersAndDomains @{Add=$user.AddSafeSender} #-Confirm $false 
                $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Sucess"",""Added to Safe Sender list: $($user.userprincipalname) : $($user.AddSafeSender)""")
            }
        }catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to Connect to Exchange Online""");$errorcount++}
        # AddBlockedSender
        try {
            if($user.AddBlockedSender -notlike "Null"){
                Set-MailboxJunkEmailConfiguration $user.userprincipalname -BlockedSendersAndDomains @{Add=$user.AddBlockedSender} -Confirm $false
                $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Sucess"",""Added to Blocked Sender list: $($user.userprincipalname) : $($user.AddBlockedSender)""")
            }
        }catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to Connect to Exchange Online""");$errorcount++}
        # RemoveSafeSender
        try {
            if($user.RemoveSafeSender -notlike "Null"){
                Set-MailboxJunkEmailConfiguration $user.userprincipalname -TrustedSendersAndDomains @{Remove=$user.RemoveSafeSender} -Confirm $false
                $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Sucess"",""Removed from SafeSender list: $($user.userprincipalname) : $($user.RemoveSafeSender)""")
            }
        }catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to Connect to Exchange Online""");$errorcount++}
        # RemoveBlockedSender
        try {
            if($user.RemoveBlockedSender -notlike "Null"){
                Set-MailboxJunkEmailConfiguration $user.userprincipalname -BlockedSendersAndDomains @{Remove=$user.RemoveBlockedSender} -Confirm $false
                $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Sucess"",""Removed from Blocked Sender list: $($user.userprincipalname) : $($user.RemoveBlockedSender)""")
            }
        }catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to Connect to Exchange Online""");$errorcount++}   
    }    
}catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to assign values to User emails""");$errorcount++}

# Disconnect from Exchange Online Session
try {
    Disconnect-ExchangeOnline -Confirm:$false
    $log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""Sucess"",""Disconnected from exchange online""")
}catch {$log.writeLine("$(get-date -uformat '%Y%m%d %H%M%S'),""ERROR"",""Failed to Disconnect from Exchange online""");$errorcount++}

# close log file
$log.Close()

# Checks if the program encountered any errors and Sends an email if it did
if ($errorCount -gt 0) {
    $anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
    $anonCredentials = New-Object System.Management.Automation.PSCredential("anonymous",$anonPassword)
    Send-MailMessage -from "Outlook_SafeSender_updates@contoso.com" -to "sysadmin@contoso.com" -Credential $anonCredentials `
    -Attachments $LogLocation -body "Error has been encountered please review attached log file." `
    -subject ((get-date -UFormat "%Y-%m-%d") + " Safe Sender update: Error") -SmtpServer "autodiscover.contoso.com"   
}