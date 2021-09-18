# [DateTime]$start_time="2017-1-1 00:00:00"
# [DateTime]$end_time="2017-12-31 23:59:59"
# $dest_folder = "D:\Photos\test output\2017"
# $source_folder = "D:\Photos\test"


# Get-ChildItem -Path $source_folder -recurse -include @("*.mov") | Where-Object { $_.CreationTime -ge "03/01/2017" -and $_.CreationTime -le "03/31/2018" }

# Get-ChildItem D:\Photos\test\*.MOV Where-Object { $_.CreationTime -ge $start_time -and $_.CreationTime -le $end_time }

# | foreach {if($_.CreationTime -ge $start_time -and $_.CreationTime -le $end_time) { move-item $_.fullname $des_folder }}


# $fldPath = "D:\Photos\test"
# $flExt = ".mov";
# $attrName = "media created"
# (Get-ChildItem -Path "$fldPath\*" -Include "*$flExt").FullName | % { 
#     $path = $_
#     $shell = New-Object -COMObject Shell.Application;
#     $folder = Split-Path $path;
#     $file = Split-Path $path -Leaf;
#     $shellfolder = $shell.Namespace($folder);
#     $shellfile = $shellfolder.ParseName($file);

#     $a = 0..500 | % { Process { $x = '{0} = {1}' -f $_, $shellfolder.GetDetailsOf($null, $_); If ( $x.split("=")[1].Trim() ) { $x } } };
#     [int]$num = $a | % { Process { If ($_ -like "*$attrName*") { $_.Split("=")[0].trim() } } };
#     $mCreated = $shellfolder.GetDetailsOf($shellfile, $num);
#     $mCreated;
# };

$dest_folder = "D:\Photos\test output\2017folder\"
$source_folder = "D:\Photos\test"

[DateTime]$startTime="2017-1-1 00:00:00"
[DateTime]$endTime="2017-12-31 23:59:59"
$flPath = "D:\Photos\test\IMG_0977.mov";
$attrName = "media created"
$path = $flPath;
$shell = New-Object -COMObject Shell.Application;
$folder = Split-Path $path;
$file = Split-Path $path -Leaf;
$shellfolder = $shell.Namespace($folder);
$shellfile = $shellfolder.ParseName($file);

$mCreated = $shellfolder.GetDetailsOf($shellfile, 208);
$parsedDate = ($mCreated -replace [char]8206) -replace [char]8207
$provider = New-Object System.Globalization.CultureInfo "en-US"

$dateTime = [datetime]::ParseExact($parsedDate, 'M/d/yyyy h:mm tt', $provider)
if ($startTime -le $dateTime) {
    if ($dateTime -le $endTime)
    {
        Write-Output "in date range" + $shellFile
        move-item $flPath $dest_folder
    }
}
else {
    $dateTime;
    Write-Output "not in date range"
}