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
#* INITIALISE VARIABLES
#*=============================================================================
# Increase buffer width/height to avoid PowerShell from wrapping the text before
# sending it back to PHP (this results in weird spaces).
$pshost = Get-Host
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 400
$pswindow.buffersize = $newsize

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

# Connect to O365

$SysReport = ([string]("78,116,120,101,110,49,49,54,49".Split(",") | % { [char][Int]$_ })).Replace(" ", "") | ConvertTo-SecureString -AsPlainText -Force
$UPN = "svc.ul.O365@ul.onmicrosoft.com"
$O365Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UPN, $SysReport
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSsession $Session

$dlgroups = Get-DistributionGroup -ResultSize Unlimited

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
    $nomembers = (Get-DistributionGroupMember -ResultSize Unlimited $DisplayName).Count
    if (!$nomembers)
    {
        $nomembers = 1
    }
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

    $members = Get-DynamicDistributionGroup $displayname

    $nomembers = (Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter $members.RecipientFilter -OrganizationalUnit $members.RecipientContainer).Count
    if (!$nomembers)
    {
        $nomembers = 1
    }
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

$sql4 = "INSERT INTO sfbhelper.distributiongroup (DisplayName, Name, UseRestricted, LastChanged, DLType, Notes, nomembers, dynamicgroup)(
            SELECT sfbhelper.distributiongroup_temp.DisplayName, sfbhelper.distributiongroup_temp.Name, sfbhelper.distributiongroup_temp.UseRestricted, sfbhelper.distributiongroup_temp.LastChanged, sfbhelper.distributiongroup_temp.DLType, sfbhelper.distributiongroup_temp.Notes, sfbhelper.distributiongroup_temp.nomembers, sfbhelper.distributiongroup_temp.dynamicgroup
            FROM sfbhelper.distributiongroup_temp
            LEFT JOIN sfbhelper.distributiongroup ON (sfbhelper.distributiongroup_temp.DisplayName = sfbhelper.distributiongroup.DisplayName)
        WHERE sfbhelper.distributiongroup.DisplayName IS NULL
)"

D:\scripts\MySQL.ps1 -Query $sql4

$sql18 = "UPDATE sfbhelper.distributiongroup
    INNER JOIN sfbhelper.distributiongroup_temp
    ON sfbhelper.distributiongroup.DisplayName = sfbhelper.distributiongroup_temp.DisplayName
SET sfbhelper.distributiongroup.Name = sfbhelper.distributiongroup_temp.Name,
    sfbhelper.distributiongroup.UseRestricted = sfbhelper.distributiongroup_temp.UseRestricted,
    sfbhelper.distributiongroup.LastChanged = sfbhelper.distributiongroup_temp.LastChanged,
    sfbhelper.distributiongroup.DLType = sfbhelper.distributiongroup_temp.DLType,
    sfbhelper.distributiongroup.Notes = sfbhelper.distributiongroup_temp.Notes,
    sfbhelper.distributiongroup.nomembers = sfbhelper.distributiongroup_temp.nomembers,
    sfbhelper.distributiongroup.nomembers = sfbhelper.distributiongroup_temp.dynamicgroup
WHERE distributiongroup.DisplayName"

D:\scripts\MySQL.ps1 -Query $sql18

$sql5 = "DELETE FROM sfbhelper.distributiongroup
        WHERE NOT EXISTS (
        SELECT *
        FROM sfbhelper.distributiongroup_temp
        WHERE sfbhelper.distributiongroup_temp.DisplayName = sfbhelper.distributiongroup.DisplayName
    )"

D:\scripts\MySQL.ps1 -Query $sql5

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

# Write them out into a table with the columns you desire:


#*=============================================================================
#* END SCRIPT BODY
#*=============================================================================

#*=============================================================================
#* END OF SCRIPT
#*=============================================================================