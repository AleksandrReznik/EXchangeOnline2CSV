<#
.SYNOPSIS
    Export mails from Exchange online mailbox folders and subfolders to single CSV file.
.DESCRIPTION     
    Powershell program which exports mails from list of folders in mailbox located in Exchange Online to single CSV file.
    Program utilize functionality provided by Microsoft Graph and does not need MS Outlook to be installed on a computer where you run the program
.PARAMETER pathToSaveFiles
    - path to create CSV file, if not specified will take current script execution path.
.PARAMETER UPN 
    - User Principal Name. Usually in a form of email of user which mailbox we going to process.
.PARAMETER maxRecordToProcessForEachFolder 
    - max records to process for each folder. If you want "no limit" put some huge value there.. like 50000000
.PARAMETER foldersToExport
    -  second level folders to start export from
.OUTPUTS
    <yyyyMMdd-HHmmss>_<UPN>_exportedGraph.csv - file created in the folder specified by pathToSaveFiles parameter:
.NOTES
    Program is created as Proof Of Concept and provided as is, with absolutely no warranty expressed or implied. Any use is at your own risk.
    Author: Aleksandr Reznik (aleksandr@reznik.lt)
#>

param (
    [string]$UPN ="user@domain.com",
    [string]$pathToSaveFiles = $PSScriptRoot +"\", #by default equals to currently run script directory   
    $maxRecordToProcessForEachFolder = 10,
    [string[]]$foldersToExport = @("Inbox","Sent Items","Deleted Items")  #second level folders to start export from
)

$global:PSOobj4CSV = @()

function Folder2PSO($folder, $parentfolderPath){
    $currFolderPath = $parentfolderPath +"\"+ $folder.DisplayName
    Write-Host "Start processing folder $($currFolderPath)"
    $childFolders = Get-MgUserMailFolderChildFolder -UserId $UPN -MailFolderId $folder.Id -All
    if ($childFolders){
        foreach($childFolder in $childFolders){
            Folder2PSO $childFolder $currFolderPath
        }
    }
    $mails = Get-MgUserMailFolderMessage -All -UserId $UPN  -MailFolderId $folder.Id
    $nrOfMails = $mails.Count
    $currMailNr = 0

    foreach($currEmail in $mails){
        $currMailNr++
        Write-Host "$($folder.DisplayName) Mail $currMailNr from $nrOfMails"
        #Write-Host "  Recepients:"
        $recipientToEmailList = $currEmail.ToRecipients.foreach{ ($_.Emailaddress) }.address -join ";"
        $recipientCCEmailList = $currEmail.CCRecipients.foreach{ ($_.Emailaddress) }.address -join ";"
        $recipientBCCEmailList = $currEmail.BCCRecipients.foreach{ ($_.Emailaddress) }.address -join ";"
        $recipientToNameList = $currEmail.ToRecipients.foreach{ ($_.Emailaddress) }.name -join ";"
        $recipientCCNameList = $currEmail.CCRecipients.foreach{ ($_.Emailaddress) }.name -join ";"
        $recipientBCCNameList = $currEmail.BCCRecipients.foreach{ ($_.Emailaddress) }.name -join ";"
        $senderEmailList = $currEmail.Sender.foreach{ ($_.Emailaddress) }.address -join ";"
        $senderNameList = $currEmail.Sender.foreach{ ($_.Emailaddress) }.name -join ";"
        
        $PSOline = [pscustomobject]@{
            'Folder' = $currFolderPath
            'ReceivedDateTime' = $currEmail.ReceivedDateTime
            'Subject' = $currEmail.Subject
            'SenderName'  = $senderNameList
            'SenderEmail'  = $senderEmailList
            'To'  = $recipientToNameList
            'ToEmails' = $recipientToEmailList
            'CC'  = $recipientCCNameList
            'CCEmails' = $recipientCCEmailList
            'BCC' = $recipientBCCNameList
            'BCCEmails' = $recipientBCCEmailList
            'Body' = $currEmail.Body.Content
        }
        $global:PSOobj4CSV += $PSOline
        If($currMailNr -gt $maxRecordToProcessForEachFolder){
            write-host "Max record limit defined in maxRecordToProcessForEachFolder varible is reached. Stopping process this folder"
            break
        }

    }
}

################################################## MAIN PROGRAM ##################################################
$moduleName = "Microsoft.Graph"
if (Get-Module -ListAvailable -Name $moduleName) {
    Write-Host "Module  $($moduleName) already installed"
} 
else {
    Write-Host "Module $($moduleName) is not intalled. Installing"
    Install-Module Microsoft.Graph -Scope CurrentUser
}
Update-Module Microsoft.Graph
Get-InstalledModule Microsoft.Graph
Import-Module Microsoft.Graph.Mail
Connect-MgGraph -Scopes "Mail.Read"
Get-MgContext # Print cuurent user's context
$folders = Get-MgUserMailFolder -UserId $UPN -All
write-host "Current folders:"
$folders.DisplayName
foreach ($currFolder in $foldersToExport){
    $currentFolder1 = $folders | Where-Object { $_.DisplayName -eq "$($currFolder)" }
    Folder2PSO $currentFolder1 ""
}
$CurrDateTimeStr=[DateTime]::Now.ToString("yyyyMMdd-HHmmss")
$pathToCSV = "$($pathToSaveFiles)$($CurrDateTimeStr)_$($UPN)_exportedGraph.csv"
$global:PSOobj4CSV|export-CSV  $pathToCSV -NoTypeInformation -append  -force
write-host "CSV is written to $pathToCSV"


