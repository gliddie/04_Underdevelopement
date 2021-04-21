param( [string] $employeeid = $(Read-Host -prompt "Please enter the EmployeeID"))


$csaduser = Get-ADObject -Properties * -Filter {(SamAccountName -eq $employeeid )} | Select-Object UserPrincipalName,DisplayName,UserAccountControl,PhysicalDeliveryOfficeName,Mail
if(-not ([string]::IsNullOrEmpty($csaduser))){
    $displayname = $csaduser.DisplayName
    $office = $csaduser.PhysicalDeliveryOfficeName
    $sipaddress = $csaduser.Mail
    $upn = "$employeeid@global.ul.com"

    if ($csaduser.UserAccountControl -match "514"){
        Write-Host "User " $displayname " is disbaled in AD. Exiting ..." -ForegroundColor Red          
        Remove-Variable * -ErrorAction SilentlyContinue
        break
    } else {
        Write-Host "User found in AD: $displayname, at office: $office"      
    }
    
} else {
    Write-Host "User not founf in AD" -ForegroundColor Red
    Remove-Variable * -ErrorAction SilentlyContinue
    break
}


Write-Host ""
Write-Host "Loooking for DID Ranges and Policies" -ForegroundColor Cyan
$policy = C:\Data\Scripts\MySQL.ps1 -Query "SELECT LocationCode FROM locationconfiguration WHERE Name LIKE '$office'"
$policy = $policy.LocationCode
if($policy){
    Write-Host "Office found in DB. OfficeCode is: $policy"
    Write-Host ""
    Write-Host "Looking for DID Ranges..." -ForegroundColor Cyan
    
    $didranges = C:\Data\Scripts\MySQL.ps1 -Query "SELECT DIDSTART,DIDEND,Notes FROM did WHERE (LocationCode LIKE '$policy') AND (SDAP LIKE '1')"
    
    if($didranges){
        Write-Host "DID Ranges found ..."        
        Write-Host ""
    }

} else {
    Write-Host "Office not found in DB." -ForegroundColor Red
    Remove-Variable * -ErrorAction SilentlyContinue

}


function Get-NewDIDMenu
{
    $index = 1
    [string]$Title = 'Please select a DID Range'
#   Clear-Host
    Write-Host "================ $Title ================"
    Write-Host ""

    foreach($did in $didranges){
        Write-Host $index": Press '$index' for DID Range:" $did.DIDSTART "Notes:" $did.Notes
        $index++
    }
    
    Write-Host "Q: Press 'Q' to quit."
}
    
if($didranges){
    Get-NewDIDMenu –Title 'My Menu'
    Write-Host ""
    $selection = Read-Host "Please make a selection"
    Write-Host "You did choose did range number:"$selection
}


if($selection){
    Write-Host "Looking for the next available DID in Range ... this can take a while ... Please wait ..."
    Write-Host ""
    if($didranges.DIDSTART.Count -gt '1'){
        $didstart = $didranges[$selection-1].DIDSTART
        $didend = $didranges[$selection-1].DIDEND
        $didwork = $didstart
    } else {
       $didstart = $didranges.DIDSTART
       $didwork = $didstart
       $didend = $didranges.DIDEND
    }

    $number1 = $didstart
    $number2 = $didend
    $worknumber = $number1
    $tabledid = $number1 -replace ".{5}$"
    $tabledid = $tabledid+'%'
    $numberrange = @()
    do{
        
        $numberrange += $worknumber
        $worknumber++
    } while($worknumber -le $number2)

    $endpointstable = C:\Data\Scripts\MySQL.ps1 -Query "SELECT DID FROM endpoints WHERE DID LIKE '$tabledid'"
    $endpointstable += C:\Data\Scripts\MySQL.ps1 -Query "SELECT DID FROM blockeddids WHERE did LIKE '$tabledid'"

    foreach($number in $numberrange){
        $etstatus = ($endpointstable.DID -contains $number)
        $number
        $etstatus
        if($etstatus -eq $false){break}
    }

    $didwork = $number
            
    # do {
    #     $didwork++
    #     Write-Host "Checking availability from DID = $didwork"
    #     $test = C:\Data\Scripts\MySQL.ps1 -Query "SELECT DID FROM endpoints WHERE DID = $didwork"
    #     $blockeddids = C:\Data\Scripts\MySQL.ps1 -Query "SELECT did FROM blockeddids WHERE did = $didwork"
    # 
    # } while ((-not ([string]::IsNullOrEmpty($test))) -or (-not ([string]::IsNullOrEmpty($blockeddids))))

    if(($didwork) -gt $didend){
        Write-Host ""
        Write-Host "DID Range Exhausted, please choose a different range or order more numbers" -ForegroundColor Yellow
        $didwork = $null
        } else {
        
                $newdid = $didwork
                Write-Host ""
                Write-Host "Sugested New DID is: +$newdid" -ForegroundColor Yellow
                Write-Host ""
        
            
    }
}

if($didranges){
    Get-NewDIDMenu –Title 'My Menu'
    Write-Host ""
    $selection = Read-Host "Please make a selection"
    Write-Host "You did choose did range number:"$selection
}