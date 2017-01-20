<#
.DESCRIPTION
Set full access on a Shared Mailbox

.PARAMETER
 
.EXAMPLE

.NOTES
Author: leif.ronnow@bestseller.com
Creation date: 20.01.2017

.CHANGELOG
Please put latest entries at the top

Author                  Date        Change
<your email address>    <date>      <change>

#>

#Parameter definition. Add your own if needed.
param(
	#-------------------------------------------------------------------------
	#Insert your parameters below
	#-------------------------------------------------------------------------
	#-------------------------------------------------------------------------
	[Parameter(Mandatory=$true)]
		[string] $SharedMailBoxName,
	[Parameter(Mandatory=$true)]
		[string] $FullAccessPersons,
	[Parameter(Mandatory=$true)]
		[string] $Requester,
	[Parameter(Mandatory=$true)]
		[string] $SRID
	#-------------------------------------------------------------------------
	#-------------------------------------------------------------------------
	#Insert your parameters above
	#-------------------------------------------------------------------------
)
$HybridWorkerJobRuntimeInfo = Get-HybridWorkerJobRuntimeInfo
$HybridWorkerConfiguration = Get-HybridWorkerConfiguration
$RunbookName = $null
$IsParent = $false
try { $RunbookName = (split-path -Path $MyInvocation.InvocationName -Leaf) -replace "....$" } catch { }
if(!$RunbookName)
{
    $IsParent = $true
	$RunbookName = $HybridWorkerJobRuntimeInfo.RunbookName
    $HybridWorkerName = $HybridWorkerConfiguration.ComputerName
    $SandboxID = $HybridWorkerJobRuntimeInfo.SandboxId
    .\Add-Log.ps1 -log_rbname $RunbookName -log_txt "HybridWorker: $HybridWorkerName"
    .\Add-Log.ps1 -log_rbname $RunbookName -log_txt "SandboxID: $SandboxID"
}

try {
    #Initializing ErrorOccurred variable, to be used throughout the runbook if needed.
    $ErrorOccurred = $false
	#-------------------------------------------------------------------------
	#Insert your code below
	#-------------------------------------------------------------------------
	#-------------------------------------------------------------------------

	
	<#
	
	*********************************
	SetFullAccessOnMailbox
	
	*********************************
	#>
  Function Set-FullAccessOnMailbox {
    Param (
      [Parameter(Mandatory=$true)] [string]$SharedMailBoxName,
      [Parameter(Mandatory=$true)] [string]$PDC,
      [Parameter(Mandatory=$true)] [string]$ExchangeServerName,
      [Parameter(Mandatory=$true)] [string]$FullAccessPersons
    )
	
    $Error.clear()
    $ContainerVariable = Get-AutomationVariable -Name 'ExchangeModifyAccount'
    $ExchangeSetFullAccess = Get-AutomationPSCredential -Name $ContainerVariable
    $ContainerVariable = Get-AutomationVariable -Name 'OUForSharedMailbox'
    $OUSharedMailbox = Get-AutomationPSCredential -Name $ContainerVariable
    $so = New-PSSessionOption -Culture 'da-DK'
    $URI = "http://$ExchangeServerName.bestcorp.net/PowerShell/"
    $ErrorOccurred = $false
    Try {
      $OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URI -Authentication Kerberos -Credential $ExchangeSetFullAccess -SessionOption $so -ErrorAction Stop # SilentlyContinue -ErrorVariable Err
      If ($OnPremSession) {
        Try {
          $Result = ''
          $ResultFromOnPrem = $null
          $ResultFromCloud = $null
          Try {
            $ResultFromOnPrem = Invoke-Command -ErrorAction Stop -Session $OnPremSession -ScriptBlock {	Param ($InvokeMailBoxName = '', $InvokeOUSharedMailbox = '');
              Get-Mailbox -Identity $InvokeMailBoxName  -OrganizationalUnit $InvokeOUSharedMailbox
            } -ArgumentList $SharedMailBoxName, $OUSharedMailbox
          }
          Catch {
            $ErrorMessage = $Error[0].Exception.Message
            .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while running the Invoke command against On Prem."
          }
          If ($ResultFromOnPrem -ne $null) {
            $Where = 'OnPrem'
          }
          Else {
            Try {
              $ResultFromCloud = Invoke-Command -ErrorAction Stop -Session $OnPremSession -ScriptBlock {	Param ($InvokeMailBoxName = '', $InvokeOUSharedMailbox = '');
                Get-RemoteMailbox -Identity $InvokeMailBoxName -OnPremisesOrganizationalUnit $InvokeOUSharedMailbox
              } -ArgumentList $SharedMailBoxName, $OUSharedMailbox
            }
            Catch {
              $ErrorMessage = $Error[0].Exception.Message
              .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while running the Invoke command against Cloud."
              $ErrorOccurred = $true
              Write-Error "BS-Error: The error $ErrorMessage occurred while running the Invoke command against Cloud."
            }
          }
          If ($ResultFromCloud -ne $null) {
            $Where = 'Cloud'
            $SharedMailBoxName = $resultfromcloud.UserPrincipalName
          }
          .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "$where"
          If ($ResultFromOnPrem -ne $null -or $ResultFromCloud -ne $null) {
            If ($Where -eq 'OnPrem') {
              $Session = $OnPremSession
            }
            If ($Where -eq 'Cloud') {
              Remove-PSSession $OnPremSession
              Try {
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $ExchangeSetFullAccess -Authentication Basic -AllowRedirection 
              }
              Catch {
                $ErrorOccurred = $true
                $ErrorMessage = $Error[0].Exception.Message
                .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while Creating PSSession."
                Write-Error "BS-Error: The error $ErrorMessage occurred while Creating PSSession."
              }
            }
          }
          Else {
            $ErrorOccurred = $true
          }
          If ($ErrorOccurred -ne $false) {
            .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while getting $SharedMailBoxName."
          }
          Else {
            .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "Mailbox $SharedMailBoxName found."
            $Result = ''
            Foreach ($FullAccessPerson in $FullAccessPersons)
            {
              Try {
                $Result = Invoke-Command -ErrorAction Stop -Session $Session -ScriptBlock {	Param ($InvokeMailBoxName = '', $InvokeFullAccessPerson = '');
                  Add-MailboxPermission -Identity $InvokeMailBoxName -User $InvokeFullAccessPerson -AccessRights FullAccess -InheritanceType All
                } -ArgumentList $SharedMailBoxName, $FullAccessPerson
              }
              Catch {
                $ErrorMessage = $Error[0].Exception.Message
                .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while running the Invoke command."
                $ErrorOccurred = $true
                Write-Error "BS-Error: The error $ErrorMessage occurred while running the Invoke command."
              }
              If ($ErrorOccurred -ne $false) {
                .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while setting Full Access for $FullAccessPerson on $SharedMailBoxName."
              }
              Else {
                .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "Fullaccess for $FullAccessPerson added on $SharedMailBoxName."
              }	
            }  
          }	
        }
        Catch {
          $ErrorMessage = $Error[0].Exception.Message
          .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while Creating PSSession against On Prem."
          $ErrorOccurred = $true
          Write-Error "BS-Error: The error $ErrorMessage occurred while Creating PSSession against On Prem."
        }
        Finally {
          #Send-MailMessage
#          If ($ErrorOccurred -eq $true) {
#            .\Send-AutomationEmail.ps1 -receiver $Requester -Subject "Failed to set Full access on $SharedMailBoxName for user $FullAccessPerson" -Message "The error `"$ErrorMessage`" occurred while setting Fullaccess on $SharedMailBoxName for $FullAccessPerson"
#          }
#          Else {
#            .\Send-AutomationEmail.ps1 -receiver $Requester -Subject 'Full access was set successfully' -Message "Full access was set on $SharedMailBoxName for person $FullAccessPerson successfully"
#          }		
          Try {
            Remove-PSSession $OnPremSession -ErrorAction Stop
          }
          Catch {
            $ErrorMessage = $Error[0].Exception.Message
            .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while Removing the PSSession."
          }
          Try {
            Remove-PSSession $Session -ErrorAction Stop
          }
          Catch {
            $ErrorMessage = $Error[0].Exception.Message
            .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while Removing the PSSession."
          }
        }	
      }
    }	
    Catch {				
      $ErrorOccurred = $true
      $ErrorMessage = $Error[0].Exception.Message
      .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "The error $ErrorMessage occurred while Creating PSSession."
      Write-Error "BS-Error: The error $ErrorMessage occurred while Creating PSSession against On Prem."
    }
    Return $ErrorOccurred
  }	
	
	<#
	*********************************
	Main
	
	
	*********************************
	#>
  $ErrorOccurred = $false
	$PDCRoleOwner = ''
	$PDCRoleOwner = .\Get-PDCRoleOwner.ps1 -log_id $log_id
	If ($PDCRoleOwner) {
		$ExchangeServerName = ''
		$ExchangeServerName = .\Get-ExchangeServer.ps1 -log_id $log_id
    If ($ExchangeServerName) {
      $FullAccessPersonsNames = $null
      $FullAccessPersonsNames = .\Get-UsernameBasedOnID.ps1 -Sys_Ids $FullAccessPersons
      
      $ReturnFromSetFullAccessOnMailbox = $false
      $ReturnFromSetFullAccessOnMailbox = Set-FullAccessOnMailbox -MailBoxName $SharedMailBoxName -PDC $PDCRoleOwner -FullAccessPerson $FullAccessPersons -ExchangeServerName $ExchangeServerName
      If ($ReturnFromSetFullAccessOnMailbox -eq $false) {
      }
      Else {
        .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt "No able to set Full Access on $SharedMailBoxName for $FullAccessPersons"
        Write-Error "BS-Error: No able to set Full Access on $SharedMailBoxName for $FullAccessPersons"
      }
    }	
    Else {
      .\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt 'Not able to connect to any Exchange server'
      Write-Error 'BS-Error: Not able to connect to any Exchange server'
    }
	}	
	Else {
		.\Add-Log.ps1 -log_id $log_id -log_rbname $RunbookName -log_txt 'No PDC found'
		Write-Error 'BS-Error: No PDC found'
	}
	#-------------------------------------------------------------------------
	#-------------------------------------------------------------------------
	#Insert your code above 
	#-------------------------------------------------------------------------
}
catch {
	$ErrorMessage = $Error[0].Exception.Message
	Write-Error "BS-Error: $ErrorMessage"
}
finally {
    if($IsParent)
    {
        $JobId = $HybridWorkerJobRuntimeInfo.JobId
        #"Is parent"
        .\Post-Runbook.ps1 -log_id $JobId
    }
}




