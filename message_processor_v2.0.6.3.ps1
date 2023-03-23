# Email Message Processor v2.0.6.3
# Written by Culpur.net
# Free to use Licensed under GNU GPLv3
# Need Help contact us msgpro [at] culpur.net
# Script last updated on 23.03.2023 at 20:27

param(
    $UserSettingsFilePath = "user_settings.json"
)

function Set-ExecutionPolicyUnrestricted {
    try {
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -ErrorAction Stop
    } catch {
        Write-Error "Failed to set execution policy: $_"
        return $false
    }
    return $true
}

function ScanOutlookFolders {
    $global:olFolders = @{
        'JVETInbox' = @{
            'Folder' = $null;
            'Processed' = $null;
            'UnProcessed' = $null;
        }
    }

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

    $jvetInbox = $inbox.Folders | Where-Object { $_.Name -eq 'JVET Inbox' }
    if ($null -eq $jvetInbox) {
        $jvetInbox = $inbox.Folders.Add('JVET Inbox')
    }
    $olFolders.JVETInbox.Folder = $jvetInbox

    $processed = $jvetInbox.Folders | Where-Object { $_.Name -eq 'Processed' }
    if ($null -eq $processed) {
        $processed = $jvetInbox.Folders.Add('Processed')
    }
    $olFolders.JVETInbox.Processed = $processed

    $unprocessed = $jvetInbox.Folders | Where-Object { $_.Name -eq 'UnProcessed' }
    if ($null -eq $unprocessed) {
        $unprocessed = $jvetInbox.Folders.Add('UnProcessed')
    }
    $olFolders.JVETInbox.UnProcessed = $unprocessed
}

function Get-UserSettings {
    $filePath = $UserSettingsFilePath
    if (Test-Path $filePath) {
        $json = Get-Content $filePath -Raw | ConvertFrom-Json
        $settings = [PSCustomObject]$json
    } else {
        $settings = @{
            SendFrom = ""
            SendTo = ""
            Smtp = ""
            Encryption = "Unencrypted"
            Interval = 2
            RI = ""
        }
    }
    Save-UserSettings
    return $settings
}

function Save-UserSettings {
    $filePath = $UserSettingsFilePath
    $json = [System.Web.Script.Serialization.JavaScriptSerializer]::new().Serialize($global:userSettings)
    Set-Content -Path $filePath -Value $json
}

$date = Get-Date -Format "ddhhmm MMM yy"
$header = @"
ENTER HEADER TEXT HERE
"@

$header2 = @"
ENTER HEADER TEXT HERE
"@

$footer = @"
ENTER FOOTER TEXT HERE
"@

function ProcessEmails {
    $folder = $global:olFolders.JVETInbox.Folder
    $emailCount = $folder.Items.Count

    Write-Host "Total messages found in JVET Inbox: $emailCount"

    for ($i = $emailCount; $i -gt 0; $i--) {
        $email = $folder.Items($i)

        if ($email.Body -match 'ENTER SOMETHING FOUND IN THE EMAIL') {
            $newEmail = CreateNewEmail -OriginalEmail $email -BeforeText $header -AfterText $footer
        } elseif ($email.Body -match 'ENTER SOMETHING FOUND IN OTHER EMAIL') {
            $newEmail = CreateNewEmail -OriginalEmail $email -BeforeText $header2 -AfterText $footer2
        }

        if ($null -ne $newEmail) {
            if (SendEmail -Email $newEmail) {
                MoveEmail -Email $email -DestinationFolder $global:olFolders.JVETInbox.Processed
                Write-Host "Successfully sent email: $($email.Subject)" -ForegroundColor Green
            } else {
                MoveEmail -Email $email -DestinationFolder $global:olFolders.JVETInbox.UnProcessed
                Write-Host "Failed to send email: $($email.Subject)" -ForegroundColor Red
            }
        }
    }

    DeleteOldProcessedEmails
}


function CreateNewEmail {
param(
$OriginalEmail,
$BeforeText,
$AfterText
)
$newEmail = $global:outlook.CreateItem(0)
$newEmail.BodyFormat = [Microsoft.Office.Interop.Outlook.OlBodyFormat]::olFormatPlain
$newEmail.Body = $BeforeText + $OriginalEmail.Body + $AfterText
$newEmail.To = $global:userSettings.SendTo
$newEmail.From = $global:userSettings.SendFrom

return $newEmail
}

function SendEmail {
param(
$Email
)
try {
$Email.Send()
return $true
} catch {
Write-Verbose "Error sending email: $_.Exception.Message" -Verbose
Add-Content -Path "errors.log" -Value "Error sending email: $_.Exception.Message"
return $false
}
}

function MoveEmail {
    param(
        $Email,
        $DestinationFolder
    )
    try {
        $Email.Move($DestinationFolder)
        Write-Host "Email moved to $($DestinationFolder.Name)" -ForegroundColor Green
    } catch {
        Write-Host "Failed to move email to $($DestinationFolder.Name): $_" -ForegroundColor Red
    }
}

function DeleteOldProcessedEmails {
    $processedFolder = $global:olFolders.JVETInbox.Processed
    $cutoffDate = (Get-Date).AddDays(-2)
    for ($i = $processedFolder.Items.Count; $i -gt 0; $i--) {
        $email = $processedFolder.Items($i)
        if ($email.ReceivedTime -lt $cutoffDate) {
            try {
                $email.Delete()
            } catch {
                Write-Host "Failed to delete email: $($email.Subject)" -ForegroundColor Red
                Write-Verbose "Error deleting email: $_" -Verbose
                Add-Content -Path "errors.log" -Value "Error deleting email: $($email.Subject)"
            }
        }
    }
}

function Main {
    Set-ExecutionPolicyUnrestricted
    ScanOutlookFolders

    $global:userSettings = Get-UserSettings -UserSettingsFilePath $UserSettingsFilePath
    $global:outlook = New-Object -ComObject Outlook.Application

    $global:userSettings.SendFrom = Read-Host -Prompt "Enter the email address you want to send from (Default: $($global:userSettings.SendFrom))" -Default $($global:userSettings.SendFrom)
    $global:userSettings.SendTo = Read-Host -Prompt "Enter the email address you want to send to (Default: $($global:userSettings.SendTo))" -Default $($global:userSettings.SendTo)
    $global:userSettings.Smtp = Read-Host -Prompt "Enter the SMTP server to use (Default: $($global:userSettings.Smtp))" -Default $($global:userSettings.Smtp)
    $global:userSettings.Encryption = Read-Host -Prompt "Select the encryption type: Unencrypted or Encrypted (Default: $($global:userSettings.Encryption))" -Default $($global:userSettings.Encryption)

    Save-UserSettings

    do {
        ProcessEmails
if ($global:userSettings.Interval) {
    Start-Sleep -Seconds ($global:userSettings.Interval * 60)
} else {
    Write-Warning "Interval not set in user settings. Using default value of 2 minutes."
    Start-Sleep -Seconds (2 * 60)
}        
#        Start-Sleep -Seconds ($global:userSettings.Interval * 60)

    } until ($Host.UI.RawUI.KeyAvailable -and ($Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode -eq 3))
}

Main
