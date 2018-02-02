#The directory needs to have \\ between the folders because this will parsed like a regular expression
param([string]$dir = 'C:\\Program Files\\VisualSVN Server', [string]$kp = $False)

#Create a report
"Report:" | Set-Content -path report.txt
function addContent {
    param([string]$defpath = "report.txt", [string]$content = "---------------------------------")
    $content | Add-Content -path $defpath
}
addContent

#Clear error array
$Error.Clear()

#Make blank arrays
$Processes = @()
$appNames = @()

#Run handle.exe to find all handles on all process in text form and store it line-by-line in array $handleOut
$handleOut = .\handle

#Go through $handleOut and find lines with processes on them that have a handle on the given directory 
foreach ($line in $handleOut) { 
    if ($line -match '\S+\spid:') {
        $exe = $line
    } elseif ($line -match $dir)  { 
        $Processes += $exe
    }
}

#Isolate the name of the process and get rid of duplicates
foreach ($process in $Processes) {
    $appNames += $process.Substring(0,($process.IndexOf(".exe")))
}
$appNames = $appNames | select -uniq

#add to log
addContent -content "Applications that have handles on ${dir}:"
foreach ($app in $appNames) {
    addContent -content "$app.exe"
}
addContent

#If enabled, attempts to end the process that has a handle on the given directory
if($kp -eq $True) {
    addContent -content "Starting process close procedure: "
    foreach ($app in $appNames) {
       Stop-Process -Name $app
       if ((get-process -name $app -ErrorAction SilentlyContinue) -ne $Null) {
            addContent -content "Process $app.exe successfully closed"
       } else {
            addContent -content "Error - Process $app.exe was not successfully closed!"
       }
    }
    addContent
}

#Error detecting
addContent -content "Errors Encountered:"

foreach ($e in $error) {
    addContent -content $e
}