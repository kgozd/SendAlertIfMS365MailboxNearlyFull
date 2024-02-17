
#Install-Module -name AzureAD
#Install-Module -Name ExchangeOnlineManagement

#connect to exchange online and azuread
function connect_ms_apps {
    param($ConfigData)

    $Username = $ConfigData.MS365Username
    $Password = ConvertTo-SecureString $ConfigData.MS365Password -AsPlainText -Force
    $UserCredential = New-Object System.Management.Automation.PSCredential ($Username, $Password)

    Connect-AzureAD -Credential $UserCredential 
    Connect-ExchangeOnline -Credential $UserCredential 
    Clear-Host
}


function close_ms_sessions {
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-AzureAD -Confirm:$false
}


function log_to_file {
    param (
        [string]$Message,
        [string]$log_file_name = "logs.txt"
    )

    $log_file_path = Join-Path -Path $PSScriptRoot -ChildPath $log_file_name

    if (-not (Test-Path $log_file_path)) {
        New-Item -Path $log_file_path -ItemType File -Force
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $formatted_message = "$timestamp - $Message"

    Add-Content -Path $log_file_path -Value $formatted_message -Force

    Write-Host $formatted_message
}


function get_user_data {
    $skuIds = @(
        "3b555118-da6a-4418-894f-7df1e2096870", # businessbasiclicense sku id
        "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46", # business premium license sku id
        "19ec0d23-8335-4cbd-94ac-6050e30712fa"  #exchangeplan 2 license sku id
    )

    $Users = @()
    foreach ($skuId in $skuIds) {
        $usersForSkuId = Get-AzureADUser -All:$true | Where-Object { $_.AssignedLicenses.SkuId -eq $skuId
        }
        $Users += $usersForSkuId
    }
  
    $UserInformation = @()
    foreach ($User in $Users) {
        try {
        $user_max_mailbox_size = (Get-Mailbox -Identity $User.UserPrincipalName ).ProhibitSendReceiveQuota.ToString()
        $user_current_mailbox_size = (Get-MailboxStatistics -Identity $User.UserPrincipalName).TotalItemSize.Value.ToString()
        }
        catch {
            log_to_file -Message "Cannot retrieve information  about $User.UserPrincipalName mailbox size"
        }
        $pattern = '(\d{1,3}(,\d{3})*(\.\d+)?)\s*bytes'
        $match_user_max_mailbox_size = [regex]::Match($user_max_mailbox_size, $pattern)
        $match_user_current_mailbox_size = [regex]::Match($user_current_mailbox_size, $pattern)

        if ($match_user_max_mailbox_size.Success -and $match_user_current_mailbox_size.Success) {
            $user_max_mailbox_size = $match_user_max_mailbox_size.Groups[1].Value -replace ',', ''
            $user_max_mailbox_size = [long]$user_max_mailbox_size
            
            $user_current_mailbox_size = $match_user_current_mailbox_size.Groups[1].Value -replace ',', ''
            $user_current_mailbox_size = [long]$user_current_mailbox_size

            $used_space_pctg = ($user_current_mailbox_size / $user_max_mailbox_size) * 100
            $used_space_pctg = [math]::Round($used_space_pctg, 2)
        }
        else {
            log_to_file -Message "Cannot define information about $User.UserPrincipalName mailbox used space"
            continue
        }

        if ($used_space_pctg -gt $ConfigData.MinimumUsedStorageToSendAlert) {
            $UserInfo = New-Object PSObject -Property @{
    
                UserDisplayName        = $User.DisplayName
                UserMail               = $User.UserPrincipalName
                MailboxUsedSpaceInPctg = $used_space_pctg 
            }
            $UserInformation += $UserInfo 
        }
    }
    return  $UserInformation
}


function send_emails {
    param ($ConfigData, $UserInformation)
    $EmailHTML = "<html><body><h2 style='font-size: 17px;'>Below is a list of mailboxes whose used space exceeds $($ConfigData.MinimumUsedStorageToSendAlert)% of the maximum mailbox space</h2>
            <table style='border-collapse: collapse; border: 2px solid black; margin: 10px; font-size: 14px; width: 80%;' cellpadding='10'><tr>
            <th style='border: 1px solid black; font-weight: bold;'>UserDisplayName</th>
            <th style='border: 1px solid black; font-weight: bold;'>UserEmailAdress</th>
            <th style='border: 1px solid black; font-weight: bold;'>MailboxUsedSpaceInPctg</th></tr>"

    $UserInformation | ForEach-Object {
        $UserInfo = $_
        $EmailHTML += "<tr><td style='border: 1px solid black;'>$($UserInfo.UserDisplayName)</td>
                <td style='border: 1px solid black;'>$($UserInfo.UserMail)</td>
                            <td style='border: 1px solid black;'>$($UserInfo.MailboxUsedSpaceInPctg)%</td></tr>"
    }
        
    $EmailHTML += "</table></body></html>"
    $EmailHTML += "<p style='color: red; font-weight: bold;'>This message has been generated automatically; please do not reply to this email!</p>"

    # Smtp server config; this below is for gmail 
    $SMTPServer = $ConfigData.SMTPServer
    $SMTPPort = $ConfigData.SMTPPort

    # credentials for email(remember for $Password variable to generate an "application secret" on gmail otherwise it wouldn't work)
    $Username = $ConfigData.SenderEmail
    $Password = ConvertTo-SecureString $ConfigData.SenderEmailPassword -AsPlainText -Force
    $SMTPCredential = New-Object System.Management.Automation.PSCredential ($Username, $Password)

    # SenderEmail is an email which will be using for automatic mail sending to every manager on list
    $EmailTo = $ConfigData.RecipientEmail
    $EmailFrom = $ConfigData.SenderEmail
    $EmailSubject = "List of Exchange mailboxes that exceed $($ConfigData.MinimumUsedStorageToSendAlert)% of the total space."

    Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -BodyAsHtml $EmailHTML -SmtpServer $SMTPServer `
        -Port $SMTPPort -UseSsl -Credential $SMTPCredential -Encoding UTF8

}


function main {
    $ConfigData = Get-Content -Path ".\CredsConfig.json" | ConvertFrom-Json

    connect_ms_apps -ConfigData $ConfigData
    $users_with__nearly_full_mailboxes = get_user_data 
    if($users_with__nearly_full_mailboxes.Count -ne 0){
        send_emails -ConfigData $ConfigData -UserInformation $users_with__nearly_full_mailboxes
    }else{
        Write-Host "None of the mailboxes meet the criteria for sending an alert" -ForegroundColor Yellow

    }
 
    close_ms_sessions
}
main
