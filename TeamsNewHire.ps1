#### Ask for the EmployeeID ####
Remove-Variable * -ErrorAction SilentlyContinue
param( [string] $employeeid = $(Read-Host -prompt "Please enter the EmployeeID"))
$mypath = $MyInvocation.MyCommand.Path
cls
##########################################
#                                        #
#        Define Variables                #
#                                        #
##########################################

Set-Variable -Name upn -Option AllScope
Set-Variable -Name csaduser -Option AllScope
Set-Variable -Name office -Option AllScope
Set-Variable -Name displayname -Option AllScope
Set-Variable -Name policy -Option AllScope
Set-Variable -Name didranges -Option AllScope
Set-Variable -Name selection -Option AllScope
Set-Variable -Name didranges -Option AllScope
Set-Variable -Name didwork -Option AllScope
Set-Variable -Name didend -Option AllScope
Set-Variable -Name selection -Option AllScope
Set-Variable -Name lineuri -Option AllScope
Set-Variable -Name newdid -Option AllScope
$status = New-Object -TypeName psobject 
Set-Variable -Name status -Option AllScope
Set-Variable -Name sipaddress -Option AllScope
Set-Variable -Name evenabled -Option AllScope

##########################################
#                                        #
#        Define Functions                #
#                                        #
##########################################

function Get-ADUserdetails
{


    Write-Host "Loading Users Details from AD" -ForegroundColor Cyan

    $csaduser = Get-ADObject -Properties * -Filter { (SamAccountName -eq $employeeid ) } | Select-Object UserPrincipalName, DisplayName, UserAccountControl, PhysicalDeliveryOfficeName, Mail, givenName
    if (-not ([string]::IsNullOrEmpty($csaduser)))
    {
        $displayname = $csaduser.DisplayName
        $office = $csaduser.PhysicalDeliveryOfficeName
        $sipaddress = $csaduser.Mail
        $upn = "$employeeid@global.ul.com"

        if ($csaduser.UserAccountControl -match "514")
        {
            Write-Host "User " $displayname " is disbaled in AD. Exiting ..." -ForegroundColor Red          
            Remove-Variable * -ErrorAction SilentlyContinue
            break
        }
        else
        {
            Write-Host "User found in AD: $displayname, at office: $office"      
        }
        
    }
    else
    {
        Write-Host "User not founf in AD" -ForegroundColor Red
        Remove-Variable * -ErrorAction SilentlyContinue
        # break
        . $mypath
    }
}

function Get-NewDIDMenu
{
    $index = 1
    [string]$Title = 'Please select a DID Range'
    #   Clear-Host
    Write-Host "================ $Title ================"
    Write-Host ""

    foreach ($did in $didranges)
    {
        Write-Host $index": Press '$index' for DID Range:" $did.DIDSTART "Notes:" $did.Notes
        $index++
    }
    
    # Write-Host "Q: Press 'Q' to quit."
}

function Check-IfEVEnabled
{
    $evenabled = C:\Data\Scripts\MySQL.ps1 -Query "SELECT TeamsEnterpriseVoiceEnabled,TeamsLineUri,DisplayName FROM endpoints WHERE EmployeeID LIKE '$employeeid'"
}

##########################################
#                                        #
#        Call Functions                  #
#                                        #
##########################################

## Load Classes

. C:\data\scripts\MyClasses.ps1

## Begin e new Object

$Teams = New-Object -TypeName Teams

## Check if user is allready EV Enabled.

Check-IfEVEnabled

## Load user settings

if ([string]::IsNullOrEmpty($evenabled.TeamsLineUri))
{
    Get-ADUserdetails
}
else
{
    $lineuri = $evenabled.TeamsLineUri
    $displayname = $evenabled.DisplayName
    Write-Host "====================== User Exists ======================"
    Write-Host ""
    Write-Host "User is allready enabled for EV and or has a Phone Number." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "User: $displayname"
    Write-Host "EmployeeID: $employeeid"
    Write-Host "LineURI: $lineuri"
    Write-Host ""
    Write-Host "====================== User Exists ======================"
    Write-Host ""
    $changenumber = Read-Host "Do you like to change the number? (y/n)."
    if ($changenumber -eq "n")
    {
        # Remove-Variable * -ErrorAction SilentlyContinue
        # break
        . $mypath
    }
}

if ($office)
{
    Write-Host "================ Choose Office ================"
    Write-Host ""
    Write-Host "Location found in AD: $office"
    Write-Host ""
    Write-Host "==============================================="
    Write-Host ""

    $officepicker = Read-Host "Would you like to keep this office? (y/n)"
    Write-Host ""
}

if ($officepicker -eq "n")
{
    
    Write-Host ""
    $officecode = Read-Host "Please enter the 3 letter office code:"
    Write-Host ""
    if ($officecode)
    {
        $Teams.policy = $officecode
        $Teams.GetSfBHelperOfficeDetails()
    }
}

if ($office)
{
    $Teams.office = $office
    $Teams.GetSfBHelperOfficeDetails()
    $didranges = $Teams.didranges
    $office = $Teams.office
}

if ($didranges)
{
    Get-NewDIDMenu –Title 'My Menu'
    Write-Host ""
    $selection = Read-Host "Please make a selection"
    Write-Host "You did choose did range number:"$selection
}

if ($selection)
{
    # Get-NewDIDs
    $Teams.didranges = $didranges
    $Teams.selection = $selection
    $Teams.DidFinder()
    $didwork = $Teams.newdid
}

if ($didwork)
{
    
    $msg = "Do you like to configure the user with these settings? (y/n)"
    
    $response = Read-Host -Prompt $msg
    if ($response -eq 'y')
    {
        $adobject = Get-ADuser -Properties * -Filter { (SAmAccountName -eq $employeeid) } | Select Mail
        $sipaddress = "sip:" + $adobject.Mail
        $lineuri = "tel:+$didwork"
        Write-Host "Making required changes in AD ..." -ForegroundColor Cyan
        Set-ADUser $employeeid -Replace @{'msRTCSIP-DeploymentLocator' = "sipfed.online.lync.com" }
        Set-ADUser $employeeid -Replace @{'msRTCSIP-FederationEnabled' = "TRUE" }
        Set-ADUser $employeeid -Replace @{'msRTCSIP-InternetAccessEnabled' = "TRUE" }
        Set-ADUser $employeeid -Replace @{'msRTCSIP-UserEnabled' = "TRUE" }
        Set-ADUser $employeeid -Replace @{'msRTCSIP-Line' = $lineuri }
        Set-ADUser $employeeid -Replace @{'msRTCSIP-PrimaryUserAddress' = $sipaddress }

        $Teams.Connect()

        Write-Host "Connecting to Teams ..." -ForegroundColor Cyan


        $Session = get-PSSession
        if ($Session -eq $null)
        { 
            Write-Host "Not connected to O365. Can't grant policies and licenses" -ForegroundColor Red
        }
        else
        {
            cls
            
            Write-Host "Granting Online Voice Routing Policy ..." -ForegroundColor Cyan
            Grant-CsOnlineVoiceRoutingPolicy $upn -PolicyName $policy
            if ($?)
            {
                Write-Host "Online Voice Routing Policy has been assigned successfull" -ForegroundColor Green
                $status | Add-Member -MemberType NoteProperty -Name voiceroutingpolicy -Value "yes"
            }
            else
            {
                Write-Host "Something went wrong. Check manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name voiceroutingpolicy -Value "no"
            }
            Write-Host "Granting Tenant Dialplan ..." -ForegroundColor Cyan
            Grant-CsTenantDialPlan $upn -PolicyName $policy
            if ($?)
            {
                Write-Host "Tenant Dial Plan has been assigned successfull" -ForegroundColor Green
                $status | Add-Member -MemberType NoteProperty -Name tenantdialplan -Value "yes"
            }
            else
            {
                Write-Host "Something went wrong. Check manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name tenantdialplan -Value "no"
            }
            Write-Host "Granting Teams Emergency Calling Policy ..." -ForegroundColor Cyan
            Grant-CsTeamsEmergencyCallingPolicy $upn -PolicyName $policy
            if ($?)
            {
                Write-Host "Teams Emergency Calling Policy has been assigned successfull" -ForegroundColor Green
                $status | Add-Member -MemberType NoteProperty -Name emergencycallingpolicy -Value "yes"
            }
            else
            {
                Write-Host "Something went wrong. Check manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name emergencycallingpolicy -Value "no"
            }
            Write-Host "Granting Teams Emergency Callrouting Policy ..." -ForegroundColor Cyan
            Grant-CsTeamsEmergencyCallRoutingPolicy $upn -PolicyName $policy
            if ($?)
            {
                Write-Host "Teams Emergency Call Routing Policy has been assigned successfull" -ForegroundColor Green
                $status | Add-Member -MemberType NoteProperty -Name emergencycallroutingpolicy -Value "yes"
            }
            else
            {
                Write-Host "Something went wrong. Check manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name emergencycallroutingpolicy -Value "no"
            }
            Write-Host "Granting TeamsOnly Mode ..." -ForegroundColor Cyan
            Grant-CsTeamsUpgradePolicy $upn -PolicyName "UpgradeToTeams"
            if ($?)
            {
                Write-Host "TeamsOnly Mode has been assigned successfull" -ForegroundColor Green
                $status | Add-Member -MemberType NoteProperty -Name teamsonly -Value "yes"
            }
            else
            {
                Write-Host "Something went wrong. Check manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name teamsonly -Value "no"
            }
            
            
                         
            Set-CsUser $upn -EnterpriseVoiceEnabled $True -HostedVoiceMail $True -ErrorAction SilentlyContinue
            
            if ($?)
            {
                C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$displayname','$employeeid','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','true','true','$policy','$policy')" 
                Write-Host "enterprisevoice has been enabled" -ForegroundColor Green
                $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "yes"
            }
            
            if ($status.enterprisevoiceenabled -ne "yes")
            {
                C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$displayname','$employeeid','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','false','false','$policy','$policy')"
                Write-Host "Enterprise voice could not be set to true. Wait one hour and then do it manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "no"
            }

            Write-Host "Assigning Phone License ..." -ForegroundColor Cyan
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicense "ul:MCOEV" -ErrorAction SilentlyContinue
            if ($?)
            {
                Write-Host "License has been assigned successfull" -ForegroundColor Green
                $status | Add-Member -MemberType NoteProperty -Name license -Value "yes"
            }
            else
            {
                Write-Host "Checking if user has any Licenses in O365" -ForegroundColor Magenta
                $license = (Get-MsolUser -UserPrincipalName $upn | Select isLicensed).IsLicensed
                if ($license -match 'False')
                {
                    Write-Host ""
                    Write-Host "User has no licenses at all." -ForegroundColor Yellow
                    Write-Host "Please wait a few hours and then try it again." -ForegroundColor Yellow
                    Write-Host "If it doesnt work by tomorrow," -ForegroundColor Yellow
                    Write-Host "please assign licemse manualy" -ForegroundColor Yellow
                    Write-Host ""
                    # Get-PSSession | Remove-PSSession
                    Remove-Variable * -ErrorAction SilentlyContinue
                    break
                }
                else
                {
                    Write-Host "User was licensed allready" -ForegroundColor Green
                    $status | Add-Member -MemberType NoteProperty -Name license -Value "yes"
                }
            }

            Clear-Host
            Write-Host "============================= Summary ============================="
            Write-Host ""
            Write-Host "Work Done So Far. These Settings have been configured:"
            Write-Host ""
            Write-Host "User                         :"$displayname
            Write-Host "EmployeeID                   :"$employeeid
            Write-Host "DID                          : +$didwork"
            Write-Host "License                      :"$status.license
            Write-Host "TenantDialPlan               :"$status.tenantdialplan
            Write-Host "VoiceRoutingPolicy           :"$status.voiceroutingpolicy
            Write-Host "EmergencyCallingPolicy       :"$status.emergencycallingpolicy
            Write-Host "EmergencyCallRoutingPolicy   :"$status.emergencycallroutingpolicy
            Write-Host "EnterpriseVoiceEnabled       :"$status.enterprisevoiceenabled
            Write-Host "TeamsOnly Mode               :"$status.teamsonly
            Write-Host ""
            Write-Host "==================================================================="

            # Get-PSSession | Remove-PSSession
            Write-Host ""
            Write-Host "An Email has been send to the User, to inform him about his new number."
            Write-Host ""
            Write-Host "Work Done. Existing ..."

            $givenname = $csaduser.givenName
            $body = "Hello $givenname ,</br></br>We have assigned telephone number +$didwork for you and enabled you for Teams Enterprise Voice.</br></br>This allows you to make and receive telephone calls. You also have a Voicemail.</br></br>Kind Regards,</br></br>UL's Unified Communications Team"

            Send-MailMessage -From 'LST.GlobalIPTAdmin@ul.com' -To $csaduser.Mail -Subject 'You have been enabled for Skype For Business Enterprise Voice' -Body $body -BodyAsHtml -SmtpServer 'smtp-relay.ul.com'
            # Remove-Variable * -ErrorAction SilentlyContinue
            . $mypath

        }
    }
    . $mypath
    
}