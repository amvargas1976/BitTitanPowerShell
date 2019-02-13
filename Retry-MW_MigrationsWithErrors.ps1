
<#

.SYNOPSIS

88888888ba   88      888888888888  88                                            88                                
88      "8b  ""    ,d     88       ""    ,d                                      88                                
88      ,8P        88     88             88                                      88                                
88aaaaaa8P'  88  MM88MMM  88       88  MM88MMM  ,adPPYYba,  8b,dPPYba,           88  8b,dPPYba,    ,adPPYba,       
88""""""8b,  88    88     88       88    88     ""     `Y8  88P'   `"8a          88  88P'   `"8a  a8"     ""       
88      `8b  88    88     88       88    88     ,adPPPPP88  88       88          88  88       88  8b               
88      a8P  88    88,    88       88    88,    88,    ,88  88       88  "88     88  88       88  "8a,   ,aa  888  
88888888P"   88    "Y888  88       88    "Y888  `"8bbdP"Y8  88       88  d8'     88  88       88   `"Ybbd8"'  888  
                                                                        8"                                         
© Copyright 2018 BitTitan, Inc. All Rights Reserved.

.DESCRIPTION
    This script needs to be run on the BitTitan Command Shell
	
.NOTES
    .Version		1.0
	Author			Antonio Vargas 
	Date			Feb/13/2019
    Disclaimer: This script is provided ‘AS IS’. No warrantee is provided either expresses or implied.
	Change Log

#>

######################################################################################################################################################
# Main Program
######################################################################################################################################################

$connectors = $null

#Working Directory
$global:workingDir = [environment]::getfolderpath("desktop")

#######################################
# Authenticate to MigrationWiz
#######################################
$creds = $host.ui.PromptForCredential("BitTitan Credentials", "Enter your BitTitan user name and password", "", "")
try {
    $mwTicket = Get-MW_Ticket -Credentials $creds
} catch {
    write-host "Error: Cannot create MigrationWiz Ticket. Error details: $($Error[0].Exception.Message)" -ForegroundColor Red
}

#######################################
# Display all document connectors
#######################################
Write-Host
Write-Host -Object  "Retrieving Document connectors ..."

Try{
    $connectors = get-mw_mailboxconnector -Ticket $mwTicket -RetrieveAll -ProjectType Storage -ErrorAction Stop
}
Catch{
    Write-Host -ForegroundColor Red -Object "ERROR: Cannot retrieve document projects."
    Exit
}

if($connectors -ne $null -and $connectors.Length -ge 1) {
    Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $connectors.Length.ToString() + " document project(s) found.") 
}
else {
    Write-Host -ForegroundColor Red -Object  "ERROR: No document projects found." 
    Exit
}

#######################################
# {Prompt for the document connector
#######################################
if($connectors -ne $null)
{
    Write-Host -ForegroundColor Yellow -Object "Select a document project:" 

    for ($i=0; $i -lt $connectors.Length; $i++)
    {
        $connector = $connectors[$i]
        Write-Host -Object $i,"-",$connector.Name,"-",$connector.ProjectType
    }
    Write-Host -Object "x - Exit"
    Write-Host

    do
    {
        $result = Read-Host -Prompt ("Select 0-" + ($connectors.Length-1) + " or x")
        if($result -eq "x")
        {
            Exit
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $connectors.Length))
        {
            $connector=$connectors[$result]
            Break
        }
    }
    while($true)

    #######################################
    # Get mailboxes
    #######################################
    $mailboxes = $null
    $MailboxesWithErrors = @()
    $MailboxErrorCount = 0
    $ExportMailboxList = @()

    Write-Host
    Write-Host -Object  ("Retrieving mailboxes for '$($connector.Name)':")

    Try{
        $mailboxes = @(Get-MW_Mailbox -Ticket $mwTicket -ConnectorId $connector.Id -RetrieveAll -ErrorAction Stop)
    }
    Catch{
        Write-Host -ForegroundColor Red "ERROR: Failed to query users in project '$($connector.Name)'"
        Exit
    }

    Foreach ($mailbox in $mailboxes){
        $LastMigration = get-MW_MailboxMigration -ticket $mwTicket -MailboxID $mailbox.id | ? {$_.Type -ne "Verification"} |Sort-Object -Property Startdate -Descending |select-object -First 1
        if ($LastMigration.Status -eq "Completed"){
            try{
                $MailboxErrors = get-mw_mailboxerror -ticket $mwTicket -mailboxid $mailbox.id -severity Error -erroraction Stop
            }
            Catch{
                Write-Host -ForegroundColor Yellow "WARNING: Cannot find errors for mailbox '$($mailbox.ExportEmailAddress)'"
            }
            if (-not ([string]::IsNullOrEmpty($MailboxErrors))){
                $MailboxesWithErrors += $mailbox
                $MailboxErrorCount = $MailboxErrorCount + $MailboxErrors.count           
            }
        }
    }

    if($MailboxesWithErrors -ne $null -and $MailboxesWithErrors.Length -ge 1)
    {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $MailboxesWithErrors.Length.ToString() + " mailbox(es) elegible to retry errors found")
        Write-Host -ForegroundColor Green -Object ("SUCCESS: '$($MailboxErrorCount)' individual errors found that will be retried")
        $RetryMigrationsSuccess = 0
        Foreach ($mailboxwitherrors in $MailboxesWithErrors){
            try{
                $RecountErrors = get-mw_mailboxerror -ticket $mwTicket -mailboxid $mailboxwitherrors.id -severity Error -erroraction Stop
                $result = Add-MW_MailboxMigration -ticket $mwTicket -mailboxid $mailboxwitherrors.id -type Repair -ConnectorId $connector.id -userid $mwTicket.userid -ErrorAction Stop
                write-host -ForegroundColor Green "INFO: Processing $($mailboxwitherrors.ExportEmailAddress) with $($RecountErrors.count) errors"
                $ErrorLine = New-Object PSCustomObject
                $ErrorLine | Add-Member -Type NoteProperty -Name MailboxID -Value $mailboxwitherrors.id
                $ErrorLine | Add-Member -Type NoteProperty -Name "Source Address" -Value $mailboxwitherrors.ExportEmailAddress
                $ErrorLine | Add-Member -Type NoteProperty -Name "Destination Address" -Value $mailboxwitherrors.ImportEmailAddress
                $ErrorLine | Add-Member -Type NoteProperty -Name "Error Count" -Value $RecountErrors.count
                $ExportMailboxList += $ErrorLine
                $RetryMigrationsSuccess = $RetryMigrationsSuccess + 1
            }
            Catch{
                Write-Host -ForegroundColor Red "ERROR: Failed to process $($mailboxwitherrors.ExportEmailAddress). Error details: $($Error[0].Exception.Message)"
            }
        }
        if ($RetryMigrationsSuccess -ge 1){
            Write-Host -ForegroundColor Yellow "INFO: $($RetryMigrationsSuccess) retry migrations executed. Exporting List to CSV."
            $ExportMailboxList | Export-CSV .\List-UsersWithErrors.csv -NoTypeInformation
        }
        Else{
            Write-Host -ForegroundColor Yellow "INFO: No retry migration passes were executed with success."
        }
    }
    else
    {
        Write-Host -ForegroundColor Yellow  "INFO: no users in project '$($connector.Name)' qualify for a retry errors pass. Make sure the users are in a completed state and have individual item errors logged."
        Exit
    }
} 


