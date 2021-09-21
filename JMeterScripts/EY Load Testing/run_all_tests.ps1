## Defaults and config

$defaultHost = 'gts-ey-upgrade'
$numberOfIterations = 1
$hostCpus = 4
$hostRam = 16
$thinkTime = "5-13"
$description = 'EY environment, document tests, with checkout fixes'

## For document load tests
$docAPIRampUpTime = 10
$docAPIDuration = 1800
$docAPIUserList = @(1,5,10,50)
$docDocumentSize = 1048576

## For document revision tests
$documentSizes = @(5242880, 10485760)
$revisions = @(1, 10, 25)
$loops = 10
$documentPrefix = ''

$runDocRevisionTests = $false
$runDocLoadTests = $true
$runMakeupTests = $false
$generateDocuments = $false

## You shouldn't need to edit anything after this, most of the time

function Get-Percentile {
    Param([int[]]$inputArray, [int]$numerator, [int]$denominator)
        
    if ($inputArray.Length % $denominator -ne 0)
    {
        return $inputArray[$numerator * ($inputArray.Length - 1) / $denominator];
    }
    $first = $inputArray[$numerator * $inputArray.Length / $denominator];
    $second = $inputArray[($numerator * $inputArray.Length /$denominator) - 1];
    return ($first + $second) / 2;
}

function Write-Meta-Info {
    Param([string]$outputDir)
    $outputFile = $outputDir + "\meta.txt"

    "Description|${description}" | Add-Content $outputFile
    "Host|${passportHost}" | Add-Content $outputFile
    "CPUs|${HostCpus}" | Add-Content $outputFile
    "RAM|${HostRam}" | Add-Content $outputFile

    $outputFile = $outputDir + "\DocumentRevision_meta.txt"

    "Document Sizes|${documentSizes} bytes" | Add-Content $outputFile
    "Rampup time|${docRampUpTime} seconds" | Add-Content $outputFile
    "Think time|${thinkTime} seconds" | Add-Content $outputFile
    "Iterations|$loops" | Add-Content $outputFile
        
    $outputFile = $outputDir + "\E2E_meta.txt"

    "Document Size|$($e2eDocumentSize/1MB)MB" | Add-Content $outputFile
    "Email Size|$($e2eEmailSize/1KB)KB" | Add-Content $outputFile
    "Rampup time|${e2eRampUpTime} seconds" | Add-Content $outputFile
    "Think time|${thinkTime} seconds" | Add-Content $outputFile
    "Duration|$($e2eDuration/60) minutes" | Add-Content $outputFile
}


function Process-Results {
    Param([string]$inputCsv,[string]$publishFolder)
    
    Write-Host "Aggregating output file $logName..."

    $csv = Import-Csv -Path $inputCsv | ? label -Match "^T\d+" | Sort label
    $grouped = $csv | Group-Object -Property label

    $resultsTable = New-Object system.Data.DataTable $inputCsv
    $resultsTable.Columns.Add("Label")
    $resultsTable.Columns.Add("Samples")
    $resultsTable.Columns["Samples"].DataType = [int]
    $resultsTable.Columns.Add("Average")
    $resultsTable.Columns["Average"].DataType = [int]
    $resultsTable.Columns.Add("Median")
    $resultsTable.Columns["Median"].DataType = [int]
    $resultsTable.Columns.Add("90% Line")
    $resultsTable.Columns["90% Line"].DataType = [int]
    $resultsTable.Columns.Add("Max")
    $resultsTable.Columns["Max"].DataType = [int]
    $resultsTable.Columns.Add("ReceivedBytes")
    $resultsTable.Columns["ReceivedBytes"].DataType = [int]
    $resultsTable.Columns.Add("SentBytes")
    $resultsTable.Columns["SentBytes"].DataType = [int]
    $resultsTable.Columns.Add("ErrorRate")
    $resultsTable.Columns["ErrorRate"].DataType = [Float]

    $grouped.ForEach({
        $elapsed = 0
        $elapsedMax = 0
        $receivedBytes = 0
        $sentBytes = 0
        $failures = 0
        $samples = ($_.Group | Measure-Object).Count
        $elapsedArray = @()

        $_.Group.ForEach({
            $elapsed += [int]$_.elapsed
            $elapsedArray += [int]$_.elapsed
            if ([int]$_.elapsed -gt $elapsedMax) {
                $elapsedMax = [int]$_.elapsed
            }
            $receivedBytes += [int]$_.bytes
            $sentBytes += [int]$_.sentBytes

            if ($_.success -eq 'false') {
                $failures += 1
            }
        })
        
        $elapsedArray = $elapsedArray | Sort-Object

        $median = Get-Percentile $elapsedArray 1 2
        $ninetyPct = Get-Percentile $elapsedArray 9 10

        $resultsTable.Rows.Add($_.Name, $samples, $elapsed/$samples, $median, $ninetyPct, $elapsedMax, $receivedBytes/$samples,
            $sentBytes/$samples, $failures/$samples)
    })

    $outputFile = "$inputCsv.aggregate.csv"

    $resultsTable | Format-Table

    $resultsTable | Export-Csv -Path $outputFile -NoTypeInformation
    
    Move-Item $inputCsv -Destination $publishFolder
    Move-Item $outputFile -Destination $publishFolder
}


$passportHost = Read-Host -Prompt "Enter the target Passport server host name (press Enter for $defaultHost)"

if ($passportHost -eq '') {
    $passportHost = $defaultHost
}

# N.B. If the user supplies bad credentials, the JMeter scripts will fail on the setup thread 
#       so that you don't waste all night waiting for junk data

$user = Read-Host -Prompt 'Enter the username to use with this host'
$pass = Read-Host -Prompt 'Enter the password to use with this user'

Write-Host "Username: $user"
Write-Host "Password: $pass" 
Write-Host "Passport URL: https://$passportHost/Passport"
Write-Host "This will take several hours to run. If you're sure, and if the above looks correct, type 'yes' to continue."
Write-Host

$confirmation = Read-Host 

if($confirmation -ne "yes"){
    Break
}

$dateString = (Get-Date).ToString('yyyy-MM-dd_HH-mm')

if (!(Test-Path "Results")) {
    New-Item -Type Directory -Name "Results"
}

$outputFolder = New-Item -Type Directory -Name "Results\$dateString"
Write-Meta-Info($outputFolder)

if ($generateDocuments) {
    Write-Host "Generating documents."

    foreach($docSize in $documentSizes) {
        foreach($revisionCount in $revisions) {
            Write-Host "Generating $docSize byte document with $revisionCount revisions..."
            Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t EY_Create_Document_With_Revisions.jmx",
                "-Jdocsize=${docSize}","-Jrevisions=${revisionCount}",
                "-Jusername=$user","-Jpassword=$pass","-Jhostname=$passportHost" -Wait -NoNewWindow
        }    
    }
}

if ($runDocRevisionTests) {
    Write-Host "Executing document revision tests."

    foreach($docSize in $documentSizes) {
        foreach($revisionCount in $revisions) {
            $logName = "DocumentRevision_${docSize}_bytes_${revisionCount}_revs_UserResults.jtl"
            Write-Host "Log name: $logName"
            Write-Host "Executing test for $docSize byte document with $revisionCount revisions..."
            Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t EY_Document_By_Size_And_Revisions.jmx",
                "-l $logName","-Jdocsize=${docSize}","-Jusers=1","-Jrevisions=${revisionCount}", "-Jloops=${loops}",
                "-Jusername=$user","-Jpassword=$pass","-Jhostname=$passportHost", "-Jprefix=$documentPrefix" -Wait -NoNewWindow
            
            Process-Results $logName $outputFolder 
        }
    }
}

if ($runDocLoadTests) {
    Write-Host "Executing document load tests."

    foreach($userNumber in $docAPIUserList) {
        $logName = "DocLoad_${userNumber}UserResults.jtl"
        For ($i=1; $i -le $numberOfIterations; $i++) {
            Write-Host "Executing iteration $userNumber-$i..."
            Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t EY_Document_Load.jmx",
                "-l $logName","-Jusers=${userNumber}","-Jrampup=${docAPIRampUpTime}","-Jduration=${docAPIDuration}",
                "-Jusername=$user","-Jpassword=$pass","-Jhostname=$passportHost", "-Jdocsize=$docDocumentSize" -Wait -NoNewWindow
        }
    
        Process-Results $logName $outputFolder
    }
}

# For one-off tests that didn't succeed in the loops above
if ($runMakeupTests) {
	Write-Host "Executing makeup document revision tests."

	$docSize = 2097152
	$revisionCount = 25

	$logName = "DocumentRevision_${docSize}_bytes_${revisionCount}_revs_UserResults.jtl"
	Write-Host "Log name: $logName"
	Write-Host "Executing test for $docSize byte document with $revisionCount revisions..."
	Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t EY_Document_By_Size_And_Revisions.jmx",
		"-l $logName","-Jdocsize=${docSize}","-Jusers=1","-Jrevisions=${revisionCount}", "-Jloops=${loops}",
		"-Jusername=$user","-Jpassword=$pass","-Jhostname=$passportHost" -Wait -NoNewWindow
				
	Process-Results $logName $outputFolder 

	$docSize = 10485760
	$revisionCount = 1

	$logName = "DocumentRevision_${docSize}_bytes_${revisionCount}_revs_UserResults.jtl"
	Write-Host "Log name: $logName"
	Write-Host "Executing test for $docSize byte document with $revisionCount revisions..."
	Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t EY_Document_By_Size_And_Revisions.jmx",
		"-l $logName","-Jdocsize=${docSize}","-Jusers=1","-Jrevisions=${revisionCount}", "-Jloops=${loops}",
		"-Jusername=$user","-Jpassword=$pass","-Jhostname=$passportHost" -Wait -NoNewWindow
				
	Process-Results $logName $outputFolder 
}

