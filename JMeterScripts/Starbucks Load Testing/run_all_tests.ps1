## Defaults and config

$defaultHost = 'starbucksqa.e-invoice.com'
$numberOfIterations = 1

$e2eRampUpTime = 1
$e2eDocSize = 1048576
$e2eMailSize = 102400
$e2eUserList = @(1)

$runE2ETests = $true

# Use these values when testing this script
#  $e2eDuration = 120
#  $docAPIDuration = 120 
#  $numberOfIterations = 1
#  $e2eUserList = @(1)
# 

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

if ($runE2ETests) {
    Write-Host "Executing end-to-end tests."
	
	
    foreach($userNumber in $e2eUserList) {	
		if ($userNumber -eq 1) {
			$logName = "E2E_SingleUserResults.jtl"
			Write-Host "Executing single-user iteration..."
			Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t Starbucks_OC_Single_User.jmx",
				"-l $logName","-Jrampup=${e2eRampUpTime}","-Jduration=${e2eDuration}",
				"-Jhostname=$passportHost", "-Jdocsize=$e2eDocSize", "-Jemailsize=$e2eMailSize" -Wait -NoNewWindow		
		}
		else {		
			$logName = "E2E_${userNumber}UserResults.jtl"
			For ($i=1; $i -le $numberOfIterations; $i++) {
				Write-Host "Executing iteration $userNumber-$i..."
				Start-Process -FilePath "D:\Programs\Apache\apache-jmeter-5.1.1\bin\jmeter.bat" -ArgumentList "-n","-t Starbucks_OC_Load.jmx",
					"-l $logName","-Jrampup=${e2eRampUpTime}","-Jduration=${e2eDuration}", "-Jusers=${userNumber}",
					"-Jhostname=$passportHost", "-Jdocsize=$e2eDocSize", "-Jemailsize=$e2eMailSize" -Wait -NoNewWindow
			}
		}

		Process-Results $logName $outputFolder    
	}
}
