#### Ask for the EmployeeID ####
param( [string] $employeeid = $(Read-Host -prompt "Please enter the EmployeeID"))
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

    $csaduser = Get-ADObject -Properties * -Filter { (SamAccountName -eq $employeeid ) } | Select-Object UserPrincipalName, DisplayName, UserAccountControl, PhysicalDeliveryOfficeName, Mail
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
        break
    }
}

function Get-SfBHelperOfficeDetails
{

    #### Look for Office and Settings ####
        
    Write-Host ""
    Write-Host "Loooking for DID Ranges and Policies" -ForegroundColor Cyan
    $policy = C:\Data\Scripts\MySQL.ps1 -Query "SELECT LocationCode FROM locationconfiguration WHERE Name LIKE '$office'"
    $policy = $policy.LocationCode
    if ($policy)
    {
        Write-Host "Office found in DB. OfficeCode is: $policy"
        Write-Host ""
        Write-Host "Looking for DID Ranges..." -ForegroundColor Cyan
        
        $didranges = C:\Data\Scripts\MySQL.ps1 -Query "SELECT DIDSTART,DIDEND,Notes FROM did WHERE (LocationCode LIKE '$policy') AND (SDAP LIKE '1')"
        
        if ($didranges)
        {
            Write-Host "DID Ranges found ..."        
            Write-Host ""
        }

    }
    else
    {
        Write-Host "Office not found in DB." -ForegroundColor Red
        Remove-Variable * -ErrorAction SilentlyContinue
    
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
    
    Write-Host "Q: Press 'Q' to quit."
}


function Get-NewDIDs
{
    if ($selection)
    {
        Write-Host "Looking for the next available DID in Range ... this can take a while ... Please wait ..."
        Write-Host ""
        if ($didranges.DIDSTART.Count -gt '1')
        {
            $didstart = $didranges[$selection - 1].DIDSTART
            $didend = $didranges[$selection - 1].DIDEND
            $didwork = $didstart - 1
        }
        else
        {
            $didstart = $didranges.DIDSTART
            $didwork = $didstart - 1
            $didend = $didranges.DIDEND
        }
    
        ################ old procedure ##################################################
        ## do {
        ##     $didwork++
        ##     Write-Host "Checking availability from DID = $didwork"
        ##     $test = C:\Data\Scripts\MySQL.ps1 -Query "SELECT DID FROM endpoints WHERE DID = $didwork"
        ##     $blockeddids = C:\Data\Scripts\MySQL.ps1 -Query "SELECT did FROM blockeddids WHERE did = $didwork"
        ## 
        ## } while ((-not ([string]::IsNullOrEmpty($test))) -or (-not ([string]::IsNullOrEmpty($blockeddids))))
    
        ################ new procedure ##################################################

        
        $number1 = $didstart
        $number2 = $didend
        $worknumber = $number1
        $tabledid = $number1 -replace ".{5}$"
        $tabledid = $tabledid + '%'
        $numberrange = @()
        do
        {
        
            $numberrange += $worknumber
            $worknumber++
        } while ($worknumber -le $number2)

        $numberrange = $numberrange | Sort-Object { Get-Random }

        $endpointstable = C:\Data\Scripts\MySQL.ps1 -Query "SELECT DID FROM endpoints WHERE DID LIKE '$tabledid'"
        $endpointstable += C:\Data\Scripts\MySQL.ps1 -Query "SELECT DID FROM blockeddids WHERE did LIKE '$tabledid'"

        foreach ($number in $numberrange)
        {
            $etstatus = ($endpointstable.DID -contains $number)
            if ($etstatus -eq $false) { break }
        }

        $didwork = $number
        
        if (($didwork) -gt $didend)
        {
            Write-Host ""
            Write-Host "DID Range Exhausted, please choose a different range or order more numbers" -ForegroundColor Yellow
            $didwork = $null
        }
        else
        {

            $newdid = $didwork
            Write-Host ""
            Write-Host "Sugested New DID is: +$newdid" -ForegroundColor Yellow
            Write-Host ""

            
        }
    }
}

function Set-LineURI
{

    $lineuri = "tel:+$newdid"

    if ([string]::IsNullOrEmpty($evenabled))
    {
        Write-Host "Enabling CsUser for: "$displayname
        Enable-CsUser $upn -HostingProviderProxyFqdn sipfed.online.lync.com -SipAddress "sip:$sipaddress" -ErrorAction SilentlyContinue
        Write-Host "Waiting one minute for replication ..." -ForegroundColor Cyan
        Write-Host "Please be patience ..." -ForegroundColor Cyan
        Start-Sleep -s 60
    }

    Write-Host "Setting LineURI for "$displayname" to: +"$newdid" ..."
    Set-CsUser $upn -LineURI $lineuri -ErrorAction SilentlyContinue
    if ($?)
    {
        Write-Host "Done ..." -ForegroundColor Green
        $status | Add-Member -MemberType NoteProperty -Name lineuri -Value "yes"

        ### Reserve DID in SfB Helper ###

        C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DID) VALUE ('$newdid')"
        
    }
    else
    {
        Write-Host "User not enabled. Do manually first ..." -ForegroundColor Red
        $status | Add-Member -MemberType NoteProperty -Name lineuri -Value "no"
    }
}

function Connect-O365
{
    $me = whoami
    $dir = "C:\Data\Scripts\"
    $File = "my" + ($me.Substring(($me.IndexOf("\") + 1), $me.length - ($me.IndexOf("\") + 1))).replace(".", "") + "File.xml"
    $moveCredFile = "c:\temp\" + $File
    $Global:CredFile = $dir + $File
    If (Test-Path $moveCredFile)
    {
        move-item $moveCredFile -destination $Global.CredFile
    }
    
    $Session = get-PSSession
    
    if ($Session -eq $null)
    {
        write-host "Connecting to Office 365....."
        If ((Test-path $Global:CredFile) -ne "True")
        {
            Get-Credential | Export-Clixml $Global:CredFile
        }
    
        $Global:LiveCred = Import-Clixml $Global:CredFile
    
        $MSOLSession = Connect-MsolService -Credential $Global:LiveCred
        $Session = New-CsOnlineSession -credential $Global:LiveCred
        Import-PSSession $Session
    
    }
    else
    {
    
        write-host "Session with Office 365 already exists." -ForegroundColor Yellow
        write-host ""
    }
    
    
    if ($Session -eq $null)
    {
        Write-Host "Not connected to O365. Something went wrong" -ForegroundColor Red
    } 
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

Check-IfEVEnabled

if ([string]::IsNullOrEmpty($evenabled.TeamsLineUri))
{
    Get-ADUserdetails
}
else
{
    $lineuri = $evenabled.TeamsLineUri
    $displayname = $evenabled.DisplayName
    Write-Host ""
    Write-Host "User is allready enabled for EV and or has a Phone Number. Please Check manually" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "User: $displayname"
    Write-Host "EmployeeID: $employeeid"
    Write-Host "LineURI: $lineuri"
    Write-Host ""
    Remove-Variable * -ErrorAction SilentlyContinue
    break
}

if ($office)
{
    Get-SfBHelperOfficeDetails
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
    Get-NewDIDs
}

if ($didwork)
{
    
    $msg = "Do you like to configure the user with these settings? (y/n)"
    
    $response = Read-Host -Prompt $msg
    if ($response -eq 'y')
    {
        Set-LineURI
        Connect-O365
        $Session = get-PSSession
        if ($Session -eq $null)
        { 
            Write-Host "Not connected to O365. Can't grant policies and licenses" -ForegroundColor Red
        }
        else
        {
            cls
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
                    Write-Host "Some times, our Team is receiving these tasks," -ForegroundColor Yellow
                    Write-Host "before Service Desk has assigned the licenses." -ForegroundColor Yellow
                    Write-Host "Please wait a few hours and then try it again." -ForegroundColor Yellow
                    Write-Host "If it doesnt work by tomorrow," -ForegroundColor Yellow
                    Write-Host "please contact Sandi Glazebrook. Good Luck ;-)" -ForegroundColor Yellow
                    Write-Host ""
                    Get-PSSession | Remove-PSSession
                    Remove-Variable * -ErrorAction SilentlyContinue
                    break
                }
                else
                {
                    Write-Host "Something went wrong. Check manually. (User might allready have the license)" -ForegroundColor Red
                    $status | Add-Member -MemberType NoteProperty -Name license -Value "no"
                }
            }
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
            
            ### Enable Enterprise Voice and Hosted Voicemail old way ###            
            
            # Set-CsUser $upn -EnterpriseVoiceEnabled $True -HostedVoiceMail $True -ErrorAction SilentlyContinue
            # if($?){
            #     C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$displayname','$employeeid','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','true','true','$policy','$policy')" 
            #     Write-Host "enterprisevoice has been enabled" -ForegroundColor Green
            #     $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "yes"
            # } else {
            #     Set-CsUser $upn -EnterpriseVoiceEnabled $True -HostedVoiceMail $True -ErrorAction SilentlyContinue
            #     if($?){
            #         C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$displayname','$employeeid','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','true','true','$policy','$policy')" 
            #         Write-Host "enterprisevoice has been enabled" -ForegroundColor Green
            #         $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "yes"
            #     } else {
            #         C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$displayname','$employeeid','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','false','false','$policy','$policy')"
            #         Write-Host "Enterprise voice could not be set to true. Wait one hour and then do it manually" -ForegroundColor Red
            #         $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "no"
            #     }
            # }

            ### Enable Enterprise Voice and Hosted Voicemail new way ###

            $i = 1
            do
            {
                Write-Host "Waiting one minute for replication ..." -ForegroundColor Cyan
                Write-Host "Please be patience ..." -ForegroundColor Cyan
                
                Start-Sleep -s 60
                
                Write-Host "Enabling Enterprise Voice and Hosted Voicemail." $i "try of 3"
                
                Set-CsUser $upn -EnterpriseVoiceEnabled $True -HostedVoiceMail $True -ErrorAction SilentlyContinue
                
                if ($?)
                {
                    $i = 10
                    C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$displayname','$employeeid','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','true','true','$policy','$policy')" 
                    Write-Host "enterprisevoice has been enabled" -ForegroundColor Green
                    $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "yes"
                }
                $i++

            } while ($i -lt '4')
            
            if ($status.enterprisevoiceenabled -ne "yes")
            {
                C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$displayname','$employeeid','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','false','false','$policy','$policy')"
                Write-Host "Enterprise voice could not be set to true. Wait one hour and then do it manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "no"
            }

            Clear-Host
            Write-Host "============================= Summary ============================="
            Write-Host ""
            Write-Host "Work Done So Far. These Settings have been configured:"
            Write-Host ""
            Write-Host "User                         :"$displayname
            Write-Host "EmployeeID                   :"$employeeid
            Write-Host "DID                          : +$newdid"
            Write-Host "License                      :"$status.license
            Write-Host "TenantDialPlan               :"$status.tenantdialplan
            Write-Host "VoiceRoutingPolicy           :"$status.voiceroutingpolicy
            Write-Host "LineURI                      :"$status.lineuri
            Write-Host "EmergencyCallingPolicy       :"$status.emergencycallingpolicy
            Write-Host "EmergencyCallRoutingPolicy   :"$status.emergencycallroutingpolicy
            Write-Host "EnterpriseVoiceEnabled       :"$status.enterprisevoiceenabled
            Write-Host "TeamsOnly Mode               :"$status.teamsonly
            Write-Host ""
            Write-Host "==================================================================="

            Get-PSSession | Remove-PSSession
            Remove-Variable * -ErrorAction SilentlyContinue
            Write-Host ""
            Write-Host "Work Done. Existing ..."
        }
    }
    
}