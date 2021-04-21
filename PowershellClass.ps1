Class Person
{
#region Properties
[string]$Name
[int] $Age
[string]$Office



#endregion Properties



#region Methods
[Void] SaySomething()
{
Write-Host "Hello from my method" -ForegroundColor Green
}



[Void] WhoAmI()
{
Write-Host "$($This.Name) - $($this.Age) - $($This.Office)" -ForegroundColor Green
}



[int] ReturnTest()
{
return $This.ReturnTest(0,100)
}



[int] ReturnTest([int]$min)
{
return $This.ReturnTest($min,100)
}



[int] ReturnTest([int]$min,[int]$max)
{
$n = Get-Random -Minimum $min -Maximum $max
return $n
}



#endregion Methods



#region Constuctor
Person()
{
$this.name = "Rick","Morty","Kory","Jeff" | Get-Random
$this.age = Get-Random -Minimum 18 -Maximum 65
$This.Office = "Home"
}



Person([string]$name,[int]$age,[string]$office)
{
$this.name = $name
$this.age = $age
$This.Office = $office
}
#endregion Constuctor



}



$MyPerson1 = [Person]::new()
$MyPerson2 = [Person]::new("Kory",31,"Homr")
$MyPerson3 = [Person]::new("Frank",35,"LA")
#[Person]::new(Kory,31,Redmond)
$MyPerson.Age = 60
$MyPerson.Name = "Rick"
$MyPerson.Office = "The Moon"