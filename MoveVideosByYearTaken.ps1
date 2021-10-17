Write-Output "***Video Mover (sort videos by Media Created)***"
$year = Read-Host -Prompt "Please enter the year to retrieve"
$fileType = Read-Host -Prompt "Please enter the desired file extension with no period (e.g. mov)"

$dest_folder = "D:\Photos\$($year) - $($fileType) files\"
$source_folder = "D:\Photos\Jeff iPhone dump 2-19-21\Camera Roll\"

[DateTime]$startTime="$($year)-1-1 00:00:00"
[DateTime]$endTime="$($year)-12-31 23:59:59"
$attrName = "media created"
$fldPath = $source_folder
$flExt = ".$($fileType)"

New-Item -ItemType Directory -Force -Path $dest_folder

(Get-ChildItem -Path "$fldPath\*" -Include "*$flExt").FullName | foreach { 

    $path = $_
    $shell = New-Object -COMObject Shell.Application;
    $folder = Split-Path $path;
    $file = Split-Path $path -Leaf;
    $shellfolder = $shell.Namespace($folder);
    $shellfile = $shellfolder.ParseName($file);

    $a = 0..500 | % { Process { $x = '{0} = {1}' -f $_, $shellfolder.GetDetailsOf($null, $_); If ( $x.split("=")[1].Trim() ) { $x } } };
    [int]$num = $a | % { Process { If ($_ -like "*$attrName*") { $_.Split("=")[0].trim() } } };
    $mCreated = $shellfolder.GetDetailsOf($shellfile, $num);
    $mCreated;
    $parsedDate = ($mCreated -replace [char]8206) -replace [char]8207
    $provider = New-Object System.Globalization.CultureInfo "en-US"

    $dateTime = [datetime]::ParseExact($parsedDate, 'M/d/yyyy h:mm tt', $provider)
    if ($startTime -le $dateTime) {
        if ($dateTime -le $endTime)
        {
            Write-Output "in date range" + $shellFile
            move-item $path $dest_folder
        }
    }
    else {
        $dateTime;
        Write-Output "not in date range"
    }
};