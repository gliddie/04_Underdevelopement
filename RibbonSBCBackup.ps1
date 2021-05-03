$sql = "SELECT hostname FROM sfbhelper.sonussbcs"
Write-Host "Loading list of Gateways from SQL DB ..." -ForegroundColor Cyan
$gws = D:\scripts\MySQL.ps1 -Query $sql
if ($?)
{
    Write-Host "Done ..." -ForegroundColor Green
} 
else 
{
    Write-Host "Gateways could not be loaded from SQL DB. Something went wrong" -ForegroundColor Red
    Break
}

foreach ($gw in $gws)
{
    $gw = $gw.hostname
    $filename = $gw + "-" + (Get-Date -format s).Replace(":", "-")
    Write-Host "Connecting to Gateway $gw ..." -ForegroundColor Cyan
    Try
    {
        connect-uxgateway -uxhostname $gw -uxusername sfbhelper -uxpassword Underwr1terS
        Write-Host "Connection to GW $gw has been established. Starting Backup ..." -ForegroundColor Green
        invoke-uxbackup -backupdestination d:\UXBackup -backupfilename $filename
    }
    Catch
    {
        Write-Host "Unable to connect to GW $gw. Skipping ..." -ForegroundColor Red
    }
}

Write-Host "Work Done" -ForegroundColor Green