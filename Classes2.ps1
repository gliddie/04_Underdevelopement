class pet
{
    #Properties
    [string] $name
    [string] $species
    #-------------
    #Constructors
    #-------------
    #Generally it looks best to keep these up near the top, after the properties are specified
    pet($n, $s)
    {
        $this.name = $n
        $this.species = $s
    }
    #-------------
    #Methods
    #-------------
    [void] DoTrick()
    {
        write-host "$($this.name) does a trick!" -ForegroundColor Green
        write-host "WOW GASP IMPRESSIVE CLAP CLAP CLAP `n" -ForegroundColor yellow
    }
}
#--------------------------------
#Now creating them is way nicer.
#You can lne break after comma separated values to make it look cleaner
#--------------------------------
$pets = [pet]::new("Mr. Whiskers", "Cat"),
[pet]::new("Doctor Barkenstein", "Dog"),
[pet]::new("Iggy", "Iguana")
#--------------------------------
foreach ($pet in $pets)
{
    $pet.DoTrick()
}

foreach ($test in $tests)
{
    Write-Host "Hallo"
}

foreach ($claus in $clauss)
{
    Write-Host ""
}




