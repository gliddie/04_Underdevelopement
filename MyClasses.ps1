## Class Connect To O365

class O365
{
    # Public Properties
    #
    # Example:
    # [String] $Name
    # [Int32]  $Age
    
    #public methods
    #
    # Example:
    # [String] SaySomething()
    # {
    #     return "Something!"
    # } 
    
    [String] Connect()
    {
        $me = whoami
        $dir = "C:\Data\Scripts\"
        $File = "my" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
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
        #    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
        
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