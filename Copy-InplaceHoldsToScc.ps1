############################################################################################# 
#  DISCLAIMER:                                                                              #
#                                                                                           #
#  THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT                #
#  PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY               #
#  OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT       #
#  LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR     #
#  PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS     #
#  AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR         #
#  ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE   # 
#  FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS  #
#  PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)   # 
#  ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,       #
#  EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES                    #
#############################################################################################
<#  
.SYNOPSIS  
    Script used to copy existing eDiscovery search from Exchange Online to Compliance Center

.DESCRIPTION  
    On launch, The srirpt connects to EXO and SCC powershell sessions. 
    The script then reads in a CSV with the names of the Exchange Online eDiscovery InPlaceHolds 
    to copy over to SCC

    For In-Place Holds it will evaluate the settings of existing holds in EXO (Get-MailboxSearch)
    and copy these settings to create the equivalent in Compliance Center

    New-ComplianceCase
    New-CaseHoldPolicy
    New-CaseHoldRule
    
.NOTES  
    File Name    : Copy-InPlaceHoldsToScc.ps1
    Requires     : Run client as admin and on powershell v3 or higher
    Author       : Joel Ricketts (joelric@microsoft.com)
#>

Param(
	[String] $InputCSV
	)

$Day = (get-date).Day.ToString()
$Month = (get-date).Month.ToString()
$Year = (get-date).Year.ToString()
$Hour = (get-date).Hour.ToString()
$Millisecond = (get-date).Millisecond.ToString()
$second = (get-date).second.ToString()
$Minute = (get-date).Minute.ToString()
$Logfilename = "CopyInPlaceHoldsToSCC-Log-File" + $Month + $Day + $Year + "-" + $Hour + $Minute + ".txt"
$Location = (get-location).Path.ToString()
$LogFilePath = "$location" + "\" + "$Filename"


Function Create-ESMLogfile #Creating and setting up Log File
{
   $DateTime = Get-Date
   Out-File -filePath $Logfilename -InputObject "$dateTime -- Strarting ESM Script......." -Append 
}

Function WriteErrorsToLog #Writes eeror to log file and exits Script
{
        #Write-Host $Error
        $DateTime = Get-Date
        Out-File -filePath $Logfilename -InputObject $DateTime -Append
        Out-File -filePath $Logfilename -InputObject "==================" -Append
        Out-File -filePath $Logfilename -InputObject $error -Append
        Out-File -filePath $Logfilename -InputObject "==================" -Append
        Out-File -filePath $Logfilename -InputObject "$dateTime -- Exiting Script......." -Append
        Write-Host $dateTime "-- Exiting Script......."
        
}

function Connect-EXOShell #prompt for credentials and connect to Exchange Online powershell
{
    Write-Host "`nPlease enter your Exchange Online admin credentials: (ie. admin@domain.onmicrosoft.com)" -ForegroundColor Green
    $global:exo_cred=Get-Credential
    Write-Host "`nConnecting to Exchange Online powershell..." -ForegroundColor Green
    Start-Sleep 1
    $exo_session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $exo_cred -Authentication Basic -AllowRedirection
	Import-PSSession $exo_session -AllowClobber -DisableNameChecking | Out-Null
}
function Connect-CompShell #connect to Compliance Center powershell, prompt for credentials optional parameter
{
    Param(
    [Parameter(Mandatory=$false,Position=1)]
    [switch]$prompt=$false
    )
    if ($prompt)
    {
        Write-Host "`nPlease enter your Security & Compliance Center admin credentials: (ie. compliance-admin@domain.onmicrosoft.com)" -ForegroundColor Green
        $global:compliance_cred=Get-Credential
        Write-Host "`nConnecting to Compliance Center powershell..." -ForegroundColor Green
        Start-Sleep 1
        $compliance_session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $compliance_cred -Authentication Basic -AllowRedirection
        Import-PSSession $compliance_session -AllowClobber -DisableNameChecking | Out-Null
    }
    else
    {
        Write-Host "`nConnecting to Security & Compliance Center powershell..." -ForegroundColor Green
        Start-Sleep 1
        $compliance_session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $compliance_cred -Authentication Basic -AllowRedirection
        Import-PSSession $compliance_session -AllowClobber -DisableNameChecking | Out-Null
    }
}


Function Import_And_Create_Legacy_EXO_InplaceHolds_In_SCC
{

    $EXOHolds = Import-Csv search.csv
    ForEach ($hold in $EXOHolds)

        {
           
            Write-Host "Sleeping for 5 seconds to help prevent service throttling" -ForegroundColor Cyan
            Start-Sleep -Seconds 5
            $DateTime = Get-Date
            $CurrentHoldName = (Get-MailboxSearch -Identity $Hold.Name -WarningAction ignore).Name.ToString()
            $CurrentWorkingHold = Get-MailboxSearch -Identity $Hold.Name -WarningAction ignore
            #$CurrentWorkingHold  | fl
            
                   

            Try {
                 $Case = New-ComplianceCase -Name $CurrentWorkingHold.Name -ErrorAction Stop
                 
                 Out-File -filePath $Logfilename -InputObject "$dateTime -- Creating New Compliance Case and Hold for $CurrentHoldName" -Append
                 Write-Host "$dateTime -- Creating New ComPlance Case and Hold for $CurrentHoldName" -ForegroundColor Green
                              

                 }

            Catch {
                  Write-Host "$dateTime -- **Unable to Create New eDiscovery Case for $CurrentHoldName" -ForegroundColor Yellow
                  Out-File -filePath $Logfilename -InputObject "$dateTime -- **Unable to Create New eDiscovery Case for $CurrentHoldName" -Append
                  Out-File -filePath $Logfilename -InputObject "$dateTime -- ** $error " -Append
                  Out-File -filePath $Logfilename -InputObject "===========================================================" -Append 
                  }
            
            Try {
                 $policy = New-CaseHoldPolicy -Name $CurrentWorkingHold.Name -Case $case.Identity -ExchangeLocation $CurrentWorkingHold.SourceMailboxes -ErrorAction Stop

                 Out-File -filePath $Logfilename -InputObject "$dateTime -- Creating New Compliance Case Hold Policy for $CurrentHoldName" -Append
                 Write-Host "$dateTime -- Creating New Compliance Case Hold Policy for $CurrentHoldName" -ForegroundColor Green
                                                      
                 }

             Catch {
                    Write-Host "$dateTime -- **Unable to Create New Case Hold Policy eDiscovery for $CurrentHoldName" -ForegroundColor Yellow
                    Out-File -filePath $Logfilename -InputObject "$dateTime -- **Unable to Create New Case Hold Policy eDiscovery for $CurrentHoldName" -Append
                    Out-File -filePath $Logfilename -InputObject "$dateTime -- ** $error " -Append 
                    Out-File -filePath $Logfilename -InputObject "===========================================================" -Append
                    
                    }

              TRy {
                    New-CaseHoldRule -Name $CurrentWorkingHold.Name -Policy $policy.Identity -ErrorAction Stop

                    Out-File -filePath $Logfilename -InputObject "$dateTime -- Creating New Compliance Case Hold Rule for $CurrentHoldName" -Append
                    Write-Host "$dateTime -- Creating New Compliance Case Hold Rule for $CurrentHoldName" -ForegroundColor Green
                    
                     }

                Catch {

                    Write-Host "$dateTime -- **Unable to Create New Case Hold Rule eDiscovery for $CurrentHoldName" -ForegroundColor Yellow
                    Out-File -filePath $Logfilename -InputObject "$dateTime -- **Unable to Create New Case Hold Rule eDiscovery for $CurrentHoldName" -Append
                    Out-File -filePath $Logfilename -InputObject "$dateTime -- ** $error " -Append 
                    Out-File -filePath $Logfilename -InputObject "===========================================================" -Append 

                }
              
            }                      
            
        
}



###################
#
# Script Begins
#
###################

# Creating Log File

Try {
    Create-ESMLogfile
    }
Catch {
    Write-Host "Cannot Create Log File" -ForegroundColor Yellow
    Write-Host $error -ForegroundColor Yellow
    exit
    }

#Connecting to EXO Powershell

Try {
        Connect-EXOShell
        $DateTime = Get-Date
        Out-File -filePath $Logfilename -InputObject "$dateTime -- Successfully Connected to Exchange Online Powershell" -Append
    }

Catch {
        WriteErrorsToLog
        exit
    }
#Connecting to SCC Powershell

Try {
        Connect-CompShell
        $DateTime = Get-Date
        Out-File -filePath $Logfilename -InputObject "$dateTime -- Successfully Connect to Secuity and Compliance Powershell" -Append
    }

Catch {
        WriteErrorsToLog
        exit
    }


## Executing Main Function

 Import_And_Create_Legacy_EXO_InplaceHolds_In_SCC

