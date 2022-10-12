$Drive="C"
$WrittenAfter = (Get-Date).AddDays(-120)  
$IR_folder = "C:\IncidentResponse\"
$fname = $IR_folder+$env:computername+"_"+$Drive+"_drive"
Set-Culture -CultureInfo 'en-AU'
(Get-Culture).DateTimeFormat.ShortDatePattern="yyyy/MM/dd"
(Get-Culture).DateTimeFormat.LongTimePattern="hh:mm:ss"
Get-ChildItem -Path $Drive":\" -Recurse -Force -ErrorAction SilentlyContinue | ? {$_.LastWriteTime -gt $WrittenAfter} | Select-Object Name,Fullname,CreationTime,LastWriteTime,Length,Attributes | Export-Csv $fname".csv"