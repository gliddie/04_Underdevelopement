#*=============================================================================
#* Script Name: ReplicateO365Details.ps1
#* Created:     2020-02-20
#* Author:  Cristian Rauth
#* Purpose:     Pull Datas from O365 and store them in a SQL Database
#*
#*=============================================================================

#*=============================================================================
#* PARAMETER DECLARATION
#*=============================================================================
param()
#*=============================================================================
#* REVISION HISTORY
#*=============================================================================
#* Date:
#* Author:
#* Purpose:
#*=============================================================================

#*=============================================================================
#* IMPORT LIBRARIES
#*=============================================================================

#*=============================================================================
#* PARAMETERS
#*=============================================================================

#*=============================================================================
#* INITIALISE VARIABLES and Classes
#*=============================================================================
# Increase buffer width/height to avoid PowerShell from wrapping the text before
# sending it back to PHP (this results in weird spaces).

. D:\Scripts\MyClasses.ps1

#*=============================================================================
#* EXCEPTION HANDLER
#*=============================================================================

#*=============================================================================
#* FUNCTION LISTINGS
#*=============================================================================

#*=============================================================================
#* Function:    Replicate Distribtion Groups
#* Created:     2020-02-20
#* Author:  Cristian Rauth
#* Purpose:
#* =============================================================================

#*=============================================================================
#* END OF FUNCTION LISTINGS
#*=============================================================================

#*=============================================================================
#* SCRIPT BODY
#*=============================================================================

# Reset Temp DBs

$sql1 = "DELETE FROM distributiongroup_temp"
$sql2 = "ALTER TABLE distributiongroup_temp AUTO_INCREMENT=1"
$sql6 = "DELETE FROM dgrp_managers_temp"
$sql7 = "ALTER TABLE dgrp_managers_temp AUTO_INCREMENT=1"
$sql11 = "DELETE FROM dgrp_authobjects_temp"
$sql12 = "ALTER TABLE dgrp_authobjects_temp AUTO_INCREMENT=1"

D:\scripts\MySQL.ps1 -Query $sql1
D:\scripts\MySQL.ps1 -Query $sql2
D:\scripts\MySQL.ps1 -Query $sql6
D:\scripts\MySQL.ps1 -Query $sql7
D:\scripts\MySQL.ps1 -Query $sql11
D:\scripts\MySQL.ps1 -Query $sql12

##############################################################
#  Connect to O365
##############################################################

$PSSession = New-Object -TypeName O365
$PSSession.Connect()

##############################################################
#  Fetch DL Groups
##############################################################

Write-Host "Fetching Distribution Groups ..." -ForegroundColor Cyan
$dlgroups = Get-DistributionGroup -ResultSize Unlimited

if ($?)
{
    Write-Host "Done. Found " $dlgroups.Count " Groups ..." -ForegroundColor Green
}
else
{
    Write-Host "Couldn't fetch Groups. Something went wrong. Fix the issue and try it again" -ForegroundColor Red
    Break
}

##############################################################
#  Prepare Data for SQL DB
##############################################################

Write-Host "Preparing fetched Datas for SQL DB ..." -ForegroundColor Cyan

foreach ($dlgroup in $dlgroups)
{
    $displayname = $dlgroup.DisplayName
    if ($displayname)
    {
        $displayname = $displayname.Replace("'", "\'")
    }
    $name = $dlgroup.Name
    if ($name)
    {
        $name = $name.Replace("'", "\'")
    }
    $userestricted = $dlgroup.RequireSenderAuthenticationEnabled
    $lastchanged = $dlgroup.WhenChanged
    $type = $dlgroup.Name.SubString(0, 3)
    $notes = $dlgroup.MailTip
    if ($notes)
    {
        $notes = $notes.Replace("'", "\'")
    }

    ##############################################################
    #  Fetching Group Members from O365
    ##############################################################

    # Write-Host "Fetchiing Grooup Members from O365 ..." -ForegroundColor Cyan
    # 
    # $nomembers = (Get-DistributionGroupMember -ResultSize Unlimited $DisplayName).Count
    # if (!$nomembers)
    # {
    #     $nomembers = 1
    # }

    ##############################################################
    #  Store collected datas in SQL Temp DB
    ##############################################################

    $nomembers = 1

    Write-Host "Store collected datas in SQL Temp DB" -ForegroundColor Cyan
    $sql3 = "INSERT INTO distributiongroup_temp (DisplayName, Name, UseRestricted, LastChanged, DLType, Notes, nomembers, dynamicgroup) VALUES ('$displayname', '$name', '$userestricted' , '$lastchanged', '$type', '$notes', '$nomembers', '0')"
    D:\scripts\MySQL.ps1 -Query $sql3

    $managers = $dlgroup.ManagedBy

    foreach ($manager in $managers)
    {
        if ($manager)
        {
            $manager = $manager.Replace("'", "\'")
        }
        $sql8 = "INSERT INTO dgrp_managers_temp (GrpName, Manager) VALUES ('$displayname', '$manager')"
        D:\scripts\MySQL.ps1 -Query $sql8
    }

    $authobjects = $dlgroup.AcceptMessagesOnlyFrom

    foreach ($authobject in $authobjects)
    {
        $sql8 = "INSERT INTO dgrp_authobjects_temp (GrpName, authobject) VALUES ('$displayname', '$authobject')"
        D:\scripts\MySQL.ps1 -Query $sql8
    }

}

##############################################################
#  Fetch dynamic groups from O365
##############################################################

Write-Host "Fetch dynamic groups from O365..." -ForegroundColor Cyan

$dlgroupsdyn = Get-DynamicDistributionGroup -ResultSize Unlimited

foreach ($dlgroupdyn in $dlgroupsdyn)
{
    $displayname = $dlgroupdyn.DisplayName
    if ($displayname)
    {
        $displayname = $displayname.Replace("'", "\'")
    }
    $name = $dlgroupdyn.Name
    if ($name)
    {
        $name = $name.Replace("'", "\'")
    }
    $userestricted = $dlgroupdyn.RequireSenderAuthenticationEnabled
    $lastchanged = $dlgroupdyn.WhenChanged
    $type = $dlgroupdyn.Name.SubString(0, 3)
    $notes = $dlgroupdyn.MailTip
    if ($notes)
    {
        $notes = $notes.Replace("'", "\'")
    }

    ##############################################################
    #  Fetch dynamic groups Members from O365
    ##############################################################
    
    # Write-Host "Fetch dynamic groups Members from O365 ..." -ForegroundColor Cyan
    # 
    # $members = Get-DynamicDistributionGroup $displayname
    # 
    # $nomembers = (Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter $members.RecipientFilter -OrganizationalUnit $members.RecipientContainer).Count
    # if (!$nomembers)
    # {
    #     $nomembers = 1
    # }
    
    $nomembersembers = 1

    ##############################################################
    #  Store Dynamic Group Details in SQL Temp DB
    ##############################################################
    
    Write-Host "Store Dynamic Group Details in SQL Temp DB" -ForegroundColor Cyan
    
    $sql15 = "INSERT INTO distributiongroup_temp (DisplayName, Name, UseRestricted, LastChanged, DLType, Notes, nomembers, dynamicgroup) VALUES ('$displayname', '$name', '$userestricted' , '$lastchanged', '$type', '$notes', '$nomembers', '1')"
    D:\scripts\MySQL.ps1 -Query $sql15

    $managers = $dlgroupdyn.ManagedBy

    foreach ($manager in $managers)
    {
        if ($manager)
        {
            $manager = $manager.Replace("'", "\'")
        }
        $sql16 = "INSERT INTO dgrp_managers_temp (GrpName, Manager) VALUES ('$displayname', '$manager')"
        D:\scripts\MySQL.ps1 -Query $sql16
    }

    $authobjects = $dlgroupdyn.AcceptMessagesOnlyFromDLMembers

    foreach ($authobject in $authobjects)
    {
        $sql17 = "INSERT INTO dgrp_authobjects_temp (GrpName, authobject) VALUES ('$displayname', '$authobject')"
        D:\scripts\MySQL.ps1 -Query $sql17
    }

}

##############################################################
#  Update production DBs with content from TempDBs...
##############################################################

Write-Host "Update production DBs with content from TempDBs" -ForegroundColor Cyan

$sql4 = "INSERT INTO sfbhelper.distributiongroup (DisplayName, Name, UseRestricted, LastChanged, DLType, Notes, nomembers, dynamicgroup)(
            SELECT sfbhelper.distributiongroup_temp.DisplayName, sfbhelper.distributiongroup_temp.Name, sfbhelper.distributiongroup_temp.UseRestricted, sfbhelper.distributiongroup_temp.LastChanged, sfbhelper.distributiongroup_temp.DLType, sfbhelper.distributiongroup_temp.Notes, sfbhelper.distributiongroup_temp.nomembers, sfbhelper.distributiongroup_temp.dynamicgroup
            FROM sfbhelper.distributiongroup_temp
            LEFT JOIN sfbhelper.distributiongroup ON (sfbhelper.distributiongroup_temp.DisplayName = sfbhelper.distributiongroup.DisplayName)
        WHERE sfbhelper.distributiongroup.DisplayName IS NULL
)"

D:\scripts\MySQL.ps1 -Query $sql4

# $sql18 = "UPDATE sfbhelper.distributiongroup
#     INNER JOIN sfbhelper.distributiongroup_temp
#     ON sfbhelper.distributiongroup.DisplayName = sfbhelper.distributiongroup_temp.DisplayName
# SET sfbhelper.distributiongroup.Name = sfbhelper.distributiongroup_temp.Name,
#     sfbhelper.distributiongroup.UseRestricted = sfbhelper.distributiongroup_temp.UseRestricted,
#     sfbhelper.distributiongroup.LastChanged = sfbhelper.distributiongroup_temp.LastChanged,
#     sfbhelper.distributiongroup.DLType = sfbhelper.distributiongroup_temp.DLType,
#     sfbhelper.distributiongroup.Notes = sfbhelper.distributiongroup_temp.Notes,
#     sfbhelper.distributiongroup.nomembers = sfbhelper.distributiongroup_temp.nomembers,
#     sfbhelper.distributiongroup.nomembers = sfbhelper.distributiongroup_temp.dynamicgroup
# WHERE distributiongroup.DisplayName"
# 
# D:\scripts\MySQL.ps1 -Query $sql18
# 
# $sql5 = "DELETE FROM sfbhelper.distributiongroup
#         WHERE NOT EXISTS (
#         SELECT *
#         FROM sfbhelper.distributiongroup_temp
#         WHERE sfbhelper.distributiongroup_temp.DisplayName = sfbhelper.distributiongroup.DisplayName
#     )"
# 
# D:\scripts\MySQL.ps1 -Query $sql5

$sql4 = "DELETE FROM distributiongroup"
$sql5 = "ALTER TABLE distributiongroup AUTO_INCREMENT=1"
D:\scripts\MySQL.ps1 -Query $sql4
D:\scripts\MySQL.ps1 -Query $sql5

# Copy Temp Table into Production Table

$sql6 = "INSERT INTO sfbhelper.distributiongroup SELECT * from sfbhelper.distributiongroup_temp"
D:\scripts\MySQL.ps1 -Query $sql6



$sql9 = "INSERT INTO sfbhelper.dgrp_managers (GrpName, Manager)(
    SELECT sfbhelper.dgrp_managers_temp.GrpName, sfbhelper.dgrp_managers_temp.Manager
    FROM sfbhelper.dgrp_managers_temp
    LEFT JOIN sfbhelper.dgrp_managers ON (sfbhelper.dgrp_managers_temp.GrpName = sfbhelper.dgrp_managers.GrpName)
    WHERE sfbhelper.dgrp_managers.GrpName IS NULL
)"

D:\scripts\MySQL.ps1 -Query $sql9

$sql10 = "DELETE FROM sfbhelper.dgrp_managers
        WHERE NOT EXISTS (
        SELECT *
        FROM sfbhelper.dgrp_managers_temp
        WHERE sfbhelper.dgrp_managers_temp.GrpName = sfbhelper.dgrp_managers.GrpName
    )"

D:\scripts\MySQL.ps1 -Query $sql10

$sql13 = "INSERT INTO sfbhelper.dgrp_authobjects (GrpName, authobject)(
    SELECT sfbhelper.dgrp_authobjects_temp.GrpName, sfbhelper.dgrp_authobjects_temp.authobject
    FROM sfbhelper.dgrp_authobjects_temp
    LEFT JOIN sfbhelper.dgrp_authobjects ON (sfbhelper.dgrp_authobjects_temp.GrpName = sfbhelper.dgrp_authobjects.GrpName)
    WHERE sfbhelper.dgrp_authobjects.GrpName IS NULL
)"

D:\scripts\MySQL.ps1 -Query $sql13

$sql14 = "DELETE FROM sfbhelper.dgrp_authobjects
        WHERE NOT EXISTS (
        SELECT *
        FROM sfbhelper.dgrp_authobjects_temp
        WHERE sfbhelper.dgrp_authobjects_temp.GrpName = sfbhelper.dgrp_authobjects.GrpName
    )"

D:\scripts\MySQL.ps1 -Query $sql14

Get-PSSession | Remove-PSSession



#*=============================================================================
#* END SCRIPT BODY
#*=============================================================================

#*=============================================================================
#* END OF SCRIPT
#*=============================================================================