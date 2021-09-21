## Defaults and config

$defaultHost = 'gts-ey-upgrade'
$numberOfIterations = 1

$documentSizes = @(1048576, 5242880, 10485760)
$revisions = @(1, 10, 25)

$prefix = 'nolocaldb'

## You shouldn't need to edit anything after this, most of the time

$passportHost = Read-Host -Prompt "Enter the target Passport server host name (press Enter for $defaultHost)"

if ($passportHost -eq '') {
    $passportHost = $defaultHost
}

# N.B. If the user supplies bad credentials, the JMeter scripts will fail on the setup thread 
#       so that you don't waste all night waiting for junk data

$user = Read-Host -Prompt 'Enter the username to use with this host'
$pass = Read-Host -Prompt 'Enter the password to use with this user'

Write-Host 
Write-Host "Username: $user"
Write-Host "Password: $pass" 
Write-Host "Passport URL: https://$passportHost/Passport"
Write-Host " If you're sure, and if the above looks correct, type 'yes' to continue."
Write-Host

$confirmation = Read-Host 

if($confirmation -ne "yes"){
    Break
}

$dateString = (Get-Date).ToString('yyyy-MM-dd_HH-mm')


$outputFolder = New-Item -Type Directory -Name "Results\$dateString"

Write-Host "Generating documents."

foreach($docSize in $documentSizes) {
    foreach($revisionCount in $revisions) {
        Write-Host "Generating $docSize byte document with $revisionCount revisions..."
        Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t EY_Create_Document_With_Revisions.jmx",
            "-Jdocsize=${docSize}","-Jrevisions=${revisionCount}",
            "-Jusername=$user","-Jpassword=$pass","-Jhostname=$passportHost", "-Jprefix=$prefix" -Wait -NoNewWindow
    }    
}

