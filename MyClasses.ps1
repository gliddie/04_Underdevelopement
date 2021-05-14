## Class Connect To O365

class O365
{
    
    [String] Connect()
    {
        $me = whoami
        $dir = "C:\Data\Scripts\"
        $File = "my" + ($me.Substring(($me.IndexOf("\") + 1), $me.length - ($me.IndexOf("\") + 1))).replace(".", "") + "File.xml"
        $moveCredFile = "c:\temp\" + $File
        $CredFile = $dir + $File
        If (Test-Path $moveCredFile)
        {
            move-item $moveCredFile -destination $CredFile
        }

        $Session = get-PSSession

        if ($Session -eq $null)
        {
            write-host "Connecting to Office 365....."
            If ((Test-path $CredFile) -ne "True")
            {
                Get-Credential | Export-Clixml $CredFile
            }
        
            $LiveCred = Import-Clixml $CredFile
        
            Connect-ExchangeOnline -Credential $LiveCred
            Add-PSSnapin *Exchange* -erroraction SilentlyContinue
            Import-Module ActiveDirectory
            Import-Module MSOnline
            Import-Module AzureAD
            Import-Module MicrosoftTeams
            write-host "Connecting to SharePointOnlineService...... " -ForegroundColor Cyan
            Connect-SPOService -url "https://ul-admin.sharepoint.com" -credential $LiveCred
            write-host "Connecting to MSOLService...... " -ForegroundColor Cyan
            Connect-MsolService -Credential $LiveCred
            write-host "Connecting to MicrosoftTeams...... " -ForegroundColor Cyan
            Connect-MicrosoftTeams -Credential $LiveCred
            write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
            Remove-PSSnapin *Exchange*
            Add-PSSnapin *Exchange*	
            Connect-AzureAD -Credential $LiveCred | Out-Null
        }
        else
        {
            write-host "Session with Office 365 already exists." -ForegroundColor Yellow
            write-host ""
        }
        return $Session
    }
    
    

}


Class Teams
{
    [int32] $employeeid
    [string] $evenabled
    [int32] $selection
    $didranges
    $newdid
    [string] $office
    [string] $policy

    [String] Connect()
    {
        $me = whoami
        $dir = "C:\Data\Scripts\"
        $File = "my" + ($me.Substring(($me.IndexOf("\") + 1), $me.length - ($me.IndexOf("\") + 1))).replace(".", "") + "File.xml"
        $moveCredFile = "c:\temp\" + $File
        $CredFile = $dir + $File

        If (Test-Path $moveCredFile)
        {
            move-item $moveCredFile -destination $CredFile
        }

        $Session = get-PSSession

        if ($Session -eq $null)
        {
            write-host "Connecting to Office 365....."
            If ((Test-path $CredFile) -ne "True")
            {
                Get-Credential | Export-Clixml $CredFile
            }
        
            $LiveCred = Import-Clixml $CredFile
        
            Import-Module ActiveDirectory
            Import-Module MicrosoftTeams
            write-host "Connecting to MicrosoftTeams...... " -ForegroundColor Cyan
            Connect-MicrosoftTeams -Credential $LiveCred
            write-host "Connecting to SkypeOnline...... " -ForegroundColor Cyan
            $Session = New-CsOnlineSession -credential $LiveCred
            Import-PSSession $Session
            C:\Users\31310.GLOBAL\"OneDrive - -"\02_PS_Scripts\01_Daily\Enable-CsOnlineSessionForReconnection.ps1
            Connect-MsolService -Credential $LiveCred
        
            #    Import-PSsession $Session -AllowClobber
        }
        else
        {
            write-host "Session with Office 365 already exists." -ForegroundColor Yellow
            write-host ""
        }
        return $Session
    }

    [String] DidFinder()
    {
        if ($this.selection)
        {
            if ($this.didranges.DIDSTART.Count -gt '1')
            {
                $didstart = $this.didranges[$this.selection - 1].DIDSTART
                $didend = $this.didranges[$this.selection - 1].DIDEND
            }
            else
            {
                $didstart = $this.didranges.DIDSTART
                $didend = $this.didranges.DIDEND
            }

            Write-Host "Looking for the next available DID in Range ... this can take a while ... Please wait ..."
            Write-Host ""

            $counter = "1"    
            $didwork = $didstart - 1
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
                cls
                Write-Host ""
                Write-Host "DID Range Exhausted, please choose a different range or order more numbers" -ForegroundColor Yellow
                $didwork = $null
            }
            else
            {
                cls
                $this.newdid = $didwork
                C:\Data\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DID) VALUES ('$didwork')"
                Write-Host ""
                Write-Host "Sugested New DID is: "+$didwork -ForegroundColor Yellow
                Write-Host ""
            

            }
        }
        else 
        {
            Write-Host "DID Ranges missing, please fix the error" -ForegroundColor Red
        }
        return $this.newdid
        
    }

    GetSfBHelperOfficeDetails()
    {

        #### Look for Office and Settings ####

        Write-Host ""
        Write-Host "Loooking for DID Ranges and Policies" -ForegroundColor Cyan
        $office2 = $this.office
        if ([string]::IsNullOrEmpty($this.policy))
        {
            $this.policy = (C:\Data\Scripts\MySQL.ps1 -Query "SELECT LocationCode FROM locationconfiguration WHERE Name LIKE '$office2'").LocationCode
            Write-Host $this.policy
        }
        if ($this.policy)
        {
            $policy2 = $this.policy
            Write-Host "Office found in DB. OfficeCode is: $policy2"
            Write-Host ""
            Write-Host "Looking for DID Ranges..." -ForegroundColor Cyan

            $this.didranges = C:\Data\Scripts\MySQL.ps1 -Query "SELECT DIDSTART,DIDEND,Notes FROM did WHERE (LocationCode LIKE '$policy2') AND (SDAP LIKE '1')"

            if ($this.didranges)
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
}

