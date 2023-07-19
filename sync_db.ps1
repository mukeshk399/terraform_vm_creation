# add path for SQLPackage.exe
IF (-not ($env:Path).Contains( "C:\Program Files\Microsoft SQL Server\160\DAC\bin"))
{ $env:path = $env:path + ";C:\Program Files\Microsoft SQL Server\160\DAC\bin;" }

sqlpackage.exe /a:extract /of:true /scs:"server=.\mydb_dev;database=db_source;trusted_connection=true" /tf:"C:\test\db_source.dacpac";

sqlpackage.exe /a:deployreport /op:"c:\test\report.xml" /of:True /sf:"C:\test\db_source.dacpac" /tcs:"server=.\sql2016; database=db_target;trusted_connection=True" 

[xml]$x = gc -Path "c:\test\report.xml";
$x.DeploymentReport.Operations.Operation |
% -Begin {$a=@();} -process {$name = $_.name; $_.Item | %  {$r = New-Object PSObject -Property @{Operation=$name; Value = $_.Value; Type = $_.Type} ; $a += $r;} }  -End {$a}
