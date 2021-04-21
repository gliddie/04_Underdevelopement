Function MyFunction
{
    [cmdletbinding()]
    param([parameter(ValueFromPipeline)]$p1,
    $p2,
    $p3)
    
    begin{write-host "starting" -ForegroundColor Magenta}
    
    
    
    process
    {
    Write-host "P1 is $p1" -ForegroundColor Green
    Write-host "P2 is $p2" -ForegroundColor Cyan
    Write-host "P3 is $p3" -ForegroundColor Yellow
    }
    
    
    
    end{write-host "ending" -ForegroundColor Magenta}
}



1,2,3 | MyFunction