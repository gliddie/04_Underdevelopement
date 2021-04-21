[cmdletbinding()]
param()
function RoundNumber
{
    param([parameter(ValueFromPipeline)]$num)
    process
    {
        [math]::Ceiling($num)
    }
}

function converto-kb
{
    param([parameter(ValueFromPipeline)]$bytes)
    process
    {
        $kb = $bytes / 1kb
        $kb = RoundNumber $kb
        write-host $kb -ForegroundColor Yellow
    }
}

1231, 3213123, 2131235 | ConverTo-KB
Write-host end -ForegroundColor Red