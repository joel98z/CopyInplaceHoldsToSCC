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

    To genrate the Input CSV file run:
    Get-MailboxSearch | ? { $_.InPlaceHoldEnabled -eq $true} | export-csv c:\temp\search.csv -NoTypeInformation

    For In-Place Holds it will evaluate the settings of existing holds in EXO (Get-MailboxSearch)
    and copy these settings to create a close equivalent in Compliance Center. 
    
    If the EXO Inplace hold has source mailboxes, the script will determine which mailboxes are active and 
    which mailboxes are inactive. Inactive Mailboxes cannot be added to a case hold policy in the compliance center
    so they will be written out to a CSV file call inactivemailbxoes+datetime.csv
    
    It is important to note that not all settings can be copied over.
        
    
    Permissions required:
    You have to be a member of the eDiscovery Manager role group in the Security & Compliance Center
    and Exchange OnlineAdministrator to run the Script


    This script is based off the guidance published at the URL Below
    https://docs.microsoft.com/en-us/microsoft-365/compliance/migrate-legacy-ediscovery-searches-and-holds?view=o365-worldwide#step-1-connect-to-exchange-online-powershell-and-security--compliance-center-powershell
    
.NOTES  
    File Name    : V3-Copy-InPlaceHoldsToScc.ps1
    Requires     : Run client as admin and on powershell v3 or higher
    Author       : Joel Ricketts (joelric@microsoft.com)
#>

Param(
                [String] $InputCSV ## Not Currently Used.. Use Search.csv as input file.
                )

$Day = (get-date).Day.ToString()
$Month = (get-date).Month.ToString()
$Year = (get-date).Year.ToString()
$Hour = (get-date).Hour.ToString()
$Millisecond = (get-date).Millisecond.ToString()
$second = (get-date).second.ToString()
$Minute = (get-date).Minute.ToString()
$Logfilename = "c:\temp\CopyInPlaceHoldsToSCC-Log-File" + $Month + $Day + $Year + "-" + $Hour + $Minute + ".txt"
#$Location = (get-location).Path.ToString()
#$LogFilePath = "$location" + "\" + "$Filename"


Function Create-ESMLogfile #Creating and setting up Log File
{
   $DateTime = Get-Date
   Out-File -filePath $Logfilename -InputObject "$dateTime -- Strarting Copy-InPlaceHoldsToScc.ps1 Script......." -Append 
}

Function WriteErrorsToLog #Writes eeror to log file and exits Script
{
        #Write-Host $Error
        $DateTime = Get-Date
        Out-File -filePath $Logfilename -InputObject $DateTime -Append
        Out-File -filePath $Logfilename -InputObject "==================" -Append
        Out-File -filePath $Logfilename -InputObject $error[0] -Append
        Out-File -filePath $Logfilename -InputObject "==================" -Append
        Out-File -filePath $Logfilename -InputObject "$dateTime -- Exiting Script......." -Append
        Write-Host $dateTime "FATAL ERROR See Log for details -- Exiting Script......."
        
}

function Connect-EXOShell #prompt for credentials and connect to Exchange Online powershell
{

    #$global:exo_cred=Get-Credential
    $livecred = Get-Credential
    $global:exo_cred=$livecred
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
    if ($prompt)  #always false - only the Else portion will execute - left here for legacy - DH
    {
        Write-Host "`nPlease enter your Security & Compliance Center admin credentials: (ie. <user>@assuranconnects.onmicrosoft.com)" -ForegroundColor Green
        $global:compliance_cred=Get-Credential
        Write-Host "`nConnecting to Compliance Center powershell..." -ForegroundColor Green
        Start-Sleep 1
        $compliance_session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $compliance_cred -Authentication Basic -AllowRedirection
        Import-PSSession $compliance_session -AllowClobber -DisableNameChecking | Out-Null
    }
    else
    {
        Write-host "Enter credentitals for Security and Compliance Center ()" -ForegroundColor Yellow
        $global:compliance_cred=get-credential
        Write-Host "`nConnecting to Security & Compliance Center powershell..." -ForegroundColor Green
        Start-Sleep 1
        $compliance_session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $compliance_cred -Authentication Basic -AllowRedirection
        Import-PSSession $compliance_session -AllowClobber -DisableNameChecking | Out-Null
    }
}


Function Import_And_Create_Legacy_EXO_InplaceHolds_In_SCC
{

    try {
       $EXOHolds = Import-Csv c:\temp\search.csv
        }
    catch {
      write-host -ForegroundColor Yellow "No Search file found - Looking for c:\temp\search.csv"
      break
          }
    
    ForEach ($hold in $EXOHolds)

        {
           
            Write-Host "Sleeping for 5 seconds to help prevent service throttling" -ForegroundColor Cyan
            Start-Sleep -Seconds 5
            $ActiveSourceMailboxCSVGileName = "c:\temp\activeSourceMailboxes.csv"
            if (Test-Path $ActiveSourceMailboxCSVGileName) {del $ActiveSourceMailboxCSVGileName}
            $ActiveSourceMailboxes = $null
            $DateTime = Get-Date
            $CurrentHoldName = (Get-MailboxSearch -Identity $Hold.Name -WarningAction ignore).Name.ToString()
            $CurrentHoldIdentity = (Get-MailboxSearch -Identity $Hold.Name -WarningAction ignore).InPlaceHoldIdentity.ToString()
            $CurrentWorkingHold = Get-MailboxSearch -Identity $Hold.Name -WarningAction ignore
            $CurrentWorkingHoldSourceMailboxes = $CurrentWorkingHold.Sources
            $CurrentWorkingHoldInactiveMailboxCSVFileName = "c:\temp\InactiveMailboxes-" + $Month + $Day + $Year + "-" + $Hour + $Minute + $second + ".csv"
            
                      
            
            # Evaluating SourceMailboxes on Legacy Hold and determing which Mailbox is an Active Mailbox and Which Mailboxes are Inactive.
            # Inactive Mailboxes cannot be added to a New Case Hold Policy. Instead they will be written out to a CSV file
            
            Write-Host "$DateTime -- Evaluating SourceMailboxes on Legacy Hold and determing which Mailbox is an Active Mailbox and Which Mailboxes are Inactive." -ForegroundColor Green
            Write-Host "$DateTime -- Currently Working on legacy In-place hold: $CurrentHoldName " -ForegroundColor Green
            
            foreach ($SourceMailboxes in $CurrentWorkingHoldSourceMailboxes ) 
                                    {
                                                                                                       
                                    $InactiveSourceMailboxes = Get-Mailbox "$SourceMailboxes" -IncludeInactiveMailbox | ? { $_.IsInactiveMailbox -eq $true } | Select PrimarySMTPAddress,Userprincipalname,ExchangeGUID,DistinguishedName,LegacyHoldName,LegacyHoldGUID,IsInactiveMailbox
                                    $ActiveSourceMailboxes = Get-Mailbox "$SourceMailboxes" -IncludeInactiveMailbox | ? { $_.IsInactiveMailbox -eq $False } | Select PrimarySMTPAddress,Userprincipalname,ExchangeGUID,DistinguishedName,IsInactiveMailbox
                                     
                                        Foreach ($Row in $InactiveSourceMailboxes) 
                                        {
                                            $row.LegacyHoldName = $CurrentHoldName
                                            $row.LegacyHoldGUID = $CurrentHoldIdentity   
                                        }
                                                                      
                                    $InactiveSourceMailboxes | export-csv $CurrentWorkingHoldInactiveMailboxCSVFileName -NoTypeInformation -Append
                                    $ActiveSourceMailboxes | export-csv $ActiveSourceMailboxCSVGileName -NoTypeInformation -Append
                                    }
                   
                   ## Creates a New Complance Case
            Try {  
                 
                 $Case = New-ComplianceCase -Name $CurrentWorkingHold.Name -ErrorAction Stop
                 
                 Out-File -filePath $Logfilename -InputObject "$dateTime -- Creating New Compliance Case for $CurrentHoldName" -Append
                 Write-Host "$dateTime -- Creating New ComPlance Case and Hold for $CurrentHoldName" -ForegroundColor Green

                 ## Setting The Description Field if the EXO Mailbox InPLaceHold had one.

                                               
                    If ($CurrentWorkingHold.Description -ne $null)
                        {
                            Set-ComplianceCase -Identity $Case.Name -Description $CurrentWorkingHold.Description -ErrorAction Stop
                            Out-File -filePath $Logfilename -InputObject "$dateTime -- Set Description for New Compliance Case $CurrentHoldName" -Append
                            Write-Host "$dateTime -- Set Description for New Compliance Case $CurrentHoldName" -ForegroundColor Green
                        }
                 }

            Catch {
                  Write-Host "$dateTime -- **Unable to Create and/or finsh creating New eDiscovery Case for $CurrentHoldName" -ForegroundColor Yellow
                  Write-Host "$dateTime -- **See Logfile for full error output" -ForegroundColor Yellow
                  Out-File -filePath $Logfilename -InputObject "$dateTime -- **Unable to Create or Configure New eDiscovery Case for $CurrentHoldName" -Append
                  Out-File -filePath $Logfilename -InputObject $error[0] -Append
                  Out-File -filePath $Logfilename -InputObject "===========================================================" -Append 
                  }

                # Creates a new Case Hold Policy for the newly created eDiscovery Case
            Try { 
                 
                 $policy = New-CaseHoldPolicy -Name $CurrentWorkingHold.Name -Case $case.Identity -ErrorAction Stop
                 Out-File -filePath $Logfilename -InputObject "$dateTime -- Creating New Compliance Case Hold Policy for $CurrentHoldName" -Append
                 Write-Host "$dateTime -- Creating New Compliance Case Hold Policy for $CurrentHoldName" -ForegroundColor Green


                 ## Setting Case Hold Policy for all Mailboxes or Specfic Mailboxes
                
                 if ($CurrentWorkingHold.AllSourceMailboxes -eq $true) 
                    {
                        Set-CaseHoldPolicy $policy.Identity -AddExchangeLocation All -ErrorAction Stop | Out-Null
                        Out-File -filePath $Logfilename -InputObject "$dateTime -- Setting Case Hold Policy for all Mailboxes" -Append
                        Write-Host "$dateTime -- Setting Case Hold Policy for all Mailboxes" -ForegroundColor Green
                    }

                        ## Setting Case Hold Policy for holds that have source mailboxes. Source Mailboxes are added one at at time.

                    Else
                        {
                           
                            $ImportedActiveSourceMailboxes = Import-CSV $ActiveSourceMailboxCSVGileName        
                            
                            Foreach ($CUrrentImportedActiveSourceMailboxes in $ImportedActiveSourceMailboxes)
                                    {
                                    $SMTPForCUrrentImportedActiveSourceMailboxes = $CUrrentImportedActiveSourceMailboxes.PrimarySmtpAddress
                                    
                                    Set-CaseHoldPolicy $policy.Identity -AddExchangeLocation "$SMTPForCUrrentImportedActiveSourceMailboxes" -ErrorAction Stop | Out-Null
                            
                                    Out-File -filePath $Logfilename -InputObject "$dateTime -- Added $SMTPForCUrrentImportedActiveSourceMailboxes to Case Hold Policy as a source mailbox" -Append
                                    Write-Host "$dateTime -- Added $SMTPForCUrrentImportedActiveSourceMailboxes to Case Hold Policy as a source mailbox" -ForegroundColor Green
                                     }

                            

                            }
                                                      
                 }

            Catch {
                    Write-Host "$dateTime -- **Unable to Create and/or finsh creating New Case Hold Policy eDiscovery for $CurrentHoldName" -ForegroundColor Yellow
                    Write-Host "$dateTime -- **See Logfile for full error output" -ForegroundColor Yellow
                    Out-File -filePath $Logfilename -InputObject "$dateTime -- **Unable to Create New Case Hold Policy eDiscovery for -- $CurrentHoldName" -Append
                    Out-File -filePath $Logfilename -InputObject $error[0] -Append 
                    Out-File -filePath $Logfilename -InputObject "===========================================================" -Append
                    
                    }

                #Creates a New Case Hold Rule for the newly created Hold POlicy. This will also set description on rule and set ContentMatchQuery if
                # There was a SearchQuery on the EXO Inplace hold.

            Try { 

                    $NewCaseHoldRule = New-CaseHoldRule -Name $CurrentWorkingHold.Name -Policy $policy.Identity -ErrorAction Stop

                    Out-File -filePath $Logfilename -InputObject "$dateTime -- Creating New Compliance Case Hold Rule for $CurrentHoldName" -Append
                    Write-Host "$dateTime -- Creating New Compliance Case Hold Rule for $CurrentHoldName" -ForegroundColor Green
                    
                         If ($CurrentWorkingHold.Description -ne $null) #### Setting The Description Field if the EXO Mailbox InPLaceHold had one
                            {

                            Set-CaseHoldRule -identity $NewCaseHoldRule.Name -Comment $CurrentWorkingHold.Description -ErrorAction Stop

                            Out-File -filePath $Logfilename -InputObject "$dateTime -- Set Description for Case Hold Rule $CurrentHoldName" -Append
                            Write-Host "$dateTime -- Set Description for Case Hold Rule $CurrentHoldName" -ForegroundColor Green
                            }


                        ElseIf ($CurrentWorkingHold.SearchQuery -ne $null) ## Setting the Content Query if the EXOMailbox Inplace hold had one
                            
                            {
                            Set-CaseHoldRule -identity $NewCaseHoldRule.Name -ContentMatchQuery $CurrentWorkingHold.SearchQuery -ErrorAction Stop

                            Out-File -filePath $Logfilename -InputObject "$dateTime -- Set ContentQuery for Case Hold Rule $CurrentHoldName" -Append
                            Write-Host "$dateTime -- Set ContentQuery for Case Hold Rule $CurrentHoldName" -ForegroundColor Green
                            }
                     }

                Catch {

                    Write-Host "$dateTime -- **Unable to Create and/or finsh creating New Case Hold Rule eDiscovery for $CurrentHoldName" -ForegroundColor Yellow
                    Write-Host "$dateTime -- **See Logfile for full error output" -ForegroundColor Yellow
                    Out-File -filePath $Logfilename -InputObject "$dateTime -- **Unable to Create New Case Hold Rule eDiscovery for $CurrentHoldName" -Append
                    Out-File -filePath $Logfilename -InputObject $error[0] -Append 
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
    break
    }

#Connecting to EXO Powershell

Try {
        Connect-EXOShell
        $DateTime = Get-Date
        Out-File -filePath $Logfilename -InputObject "$dateTime -- Successfully Connected to Exchange Online Powershell" -Append
    }

Catch {
        WriteErrorsToLog
        write-host -ForegroundColor yellow "Cannot Connect to EXOShell"
        break
     }
#Connecting to SCC Powershell

Try {
        Connect-CompShell
        $DateTime = Get-Date
        Out-File -filePath $Logfilename -InputObject "$dateTime -- Successfully Connect to Secuity and Compliance Powershell" -Append
    }

Catch {
        WriteErrorsToLog
        write-host -ForegroundColor yellow  "Cannot connect to Compshell"
        break
    }


## Executing Main Function

Import_And_Create_Legacy_EXO_InplaceHolds_In_SCC

if (Test-Path c:\temp\activeSourceMailboxes.csv) {del c:\temp\activeSourceMailboxes.csv}

$DateTime = Get-Date
Write-Host "$DateTime -- Script Completed! " -ForegroundColor Green
Out-File -filePath $Logfilename -InputObject "$DateTime -- Script Completed!" -Append


