<#

.DESCRIPTION
    This script will move mailboxes from a mailbox project to a target project
	
.NOTES
	Author			Antonio Vargas
	Date		    Jan/2019
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied.
    Version: 1.1
#>

### Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory=$true)] [string]$workingDir,
        [parameter(Mandatory=$true)] [string]$logDir
    )
    if ( !(Test-Path -Path $workingDir)) {
		try {
			$suppressOutput = New-Item -ItemType Directory -Path $workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
		}
		catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
		}
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

### Function to write information to the Log File
Function Log-Write
{
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}

### Function to display the workgroups created by the user
Function Select-MSPC_Workgroup {

    #######################################
    # Display all mailbox workgroups
    #######################################

    $workgroupPageSize = 100
  	$workgroupOffSet = 0
	$workgroups = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC workgroups ..."

    do
    {
        $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize)
    
        if($workgroupsPage) {
            $workgroups += @($workgroupsPage)
            foreach($Workgroup in $workgroupsPage) {
                Write-Progress -Activity ("Retrieving workgroups (" + $workgroups.Length + ")") -Status $Workgroup.Id
            }

            $workgroupOffset += $workgroupPageSize
        }

    } while($workgroupsPage)

    if($workgroups -ne $null -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $workgroups.Length.ToString() + " Workgroup(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }

    #######################################
    # Prompt for the mailbox Workgroup
    #######################################
    if($workgroups -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -ForegroundColor Gray -Object "INFO: your default workgroup has no name, only Id." 

        for ($i=0; $i -lt $workgroups.Length; $i++)
        {
            $Workgroup = $workgroups[$i]
            if($Workgroup.Name -eq $null) {
                Write-Host -Object $i,"-",$Workgroup.Id
            }
            else {
                Write-Host -Object $i,"-",$Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($workgroups.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0 or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($workgroups.Length-1) + ", or x")
            }
            
            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length))
            {
                $Workgroup=$workgroups[$result]
                Return $Workgroup.Id
            }
        }
        while($true)

    }

}

### Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$WorkgroupId
    )

    #######################################
    # Display all mailbox customers
    #######################################

    $customerPageSize = 100
  	$customerOffSet = 0
	$customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers ..."

    do
    {
        $customersPage = @(Get-BT_Customer -WorkgroupId $WorkgroupId -IsDeleted False -IsArchived False -PageOffset $customerOffSet -PageSize $customerPageSize)
    
        if($customersPage) {
            $customers += @($customersPage)
            foreach($customer in $customersPage) {
                Write-Progress -Activity ("Retrieving customers (" + $customers.Length + ")") -Status $customer.CompanyName
            }

            $customerOffset += $customerPageSize
        }

    } while($customersPage)

    if($customers -ne $null -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $customers.Length.ToString() + " customer(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox customer
    #######################################
    if($customers -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a customer:" 

        for ($i=0; $i -lt $customers.Length; $i++)
        {
            $customer = $customers[$i]
            Write-Host -Object $i,"-",$customer.CompanyName
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($customers.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0 or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length-1) + ", or x")
            }

            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length))
            {
                $customer=$customers[$result]
                Return $Customer.OrganizationId
            }
        }
        while($true)

    }

}

### Function to display all mailbox connectors
Function Select-MW_Connector {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerId
    )

    #######################################
    # Display all mailbox connectors
    #######################################
    
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving mailbox connectors ..."
    
    do
    {
        $connectorsPage = @(Get-MW_MailboxConnector -ticket $global:mwTicket -OrganizationId $customerId -PageOffset $connectorOffSet -PageSize $connectorPageSize)
    
        if($connectorsPage) {
            $connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($connectors -ne $null -and $connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $connectors.Length.ToString() + " mailbox connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No mailbox connectors found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox connector
    #######################################
    if($connectors -ne $null)
    {
        

        for ($i=0; $i -lt $connectors.Length; $i++)
        {
            $connector = $connectors[$i]
            Write-Host -Object $i,"-",$connector.Name
        }
        Write-Host -Object "x - Exit"
        Write-Host

        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the source mailbox connector:" 

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($connectors.Length-1) + " o x")
            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $connectors.Length))
            {
                $global:connector=$connectors[$result]
                Break
            }
        }
        while($true)
    }

}

Function Add-MW_Category {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Connector
    )

    # add items to a MigrationWiz project

    $count = 0

    Write-Host                                                                   
    Write-Host -Object  ("Aplying categories to migration item(s) in the MigrationWiz project " + $connector.Name)
	$importFilename = (Read-Host -prompt "Enter the full path to CSV import file")

	# read csv file
	$users = Import-Csv -Path $importFilename
	foreach($user in $users)
	{
	    $sourceEmail = $user.'Source Email'
        $flags = $user.'Flags'

		if($sourceEmail -ne $null -and $sourceEmail -ne "" -and $flags -in 1..6)
		{
            $count++
            Write-Progress -Activity ("Applying category to migration item (" + $count + ")") -Status $sourceEmail
            $mbx = get-mw_mailbox -ticket $mwTicket -ExportEmailAddress $sourceEmail
            if ($mbx)
            {
                $Category = ";tag-"+$flags+";"
                $result = Set-MW_Mailbox -Ticket $mwTicket -ConnectorId $connector.Id -mailbox $mbx -Categories $Category
            }
            else 
            {
                Write-Host "Cannot find MigrationWiz line item with source address: '$($sourceEmail)'" -ForegroundColor Yellow  
            }
        }
        else {
            Write-Host "The line item with the address '$($sourceEmail)' and the flag '$($flags)' is not valid." -ForegroundColor Yellow
        }
	}
    
    if($count -eq 1)
    {
        Write-Host -Object "1 mailbox has been categorized in",$connector.Name -ForegroundColor Green
    }
    if($count -ge 2)
    {
        Write-Host -Object $count," mailboxes have been categorized in",$connector.Name -ForegroundColor Green
    }

}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Move-MW_Mailboxes.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
try {
    # Get a ticket and set it as default
    $ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan -SetDefault
    # Get a MW ticket
    $global:mwTicket = Get-MW_Ticket -Credentials $creds 
} catch {
    $msg = "ERROR: Failed to create ticket."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message    

    Exit
}

#Select workgroup
$WorkgroupId = Select-MSPC_WorkGroup

#Select customer
$customerId = Select-MSPC_Customer -Workgroup $WorkgroupId

#Select connector
Select-MW_Connector -customerId $customerId 
$result = Add-MW_Category -Connector $connector


$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT