$date=get-date
$dateFormatted=get-date -format yyyyMMdd.HHmmss
$filename="RDConnectionReports_$dateFormatted.csv"
$lastCapture=([Environment]::GetEnvironmentVariable("RDLogsCollected","Machine"))
$logs=get-winevent -provider "Microsoft-Windows-TerminalServices-LocalSessionManager" | where {($_.id -eq "21"-or $_.id -eq "23" -or $_.id -eq "24") -and ($_.TimeCreated -gt $lastCapture)}

$data=@()

foreach($line in $logs){
	
	$time=$line.TimeCreated
	$eventId=$line.iD
	switch ($eventId){
		21 {$event="logon"}
		23 {$event="logoff"}
		24 {$event="disconnect"}
	}

	$message=$line.message
	
	$aliasIndexStart=($message.indexof("User:")) + 6
	$aliasIndexEnd=$message.indexof("Session ID:") - $aliasIndexStart - 2
    $alias=($message.substring($aliasIndexStart,$aliasIndexEnd)).toUpper() -replace "idi-hq\\",""
    
    if($message -like "*Source Network Address*"){
        	$sessionIDIndexStart=($message.indexof("Session ID:")) + 12
	        $sessionIDIndexEnd=$message.indexof("Source Network Address:") - $sessionIDIndexStart - 2
            $sessionID=($message.substring($sessionIDIndexStart,$sessionIDIndexEnd))
    }
    else{
        $sessionIDIndexStart=($message.indexof("Session ID:")) + 12
        $sessionID=($message.substring($sessionIDIndexStart))
    }


    if($message -like "*Source Network Address*"){
        $sourceIPIndexStart=($message.indexof("Source Network Address:")) + 24
        $sourceIP=($message.substring($sourceIPIndexStart))
    }
    else{
        $sourceIP=$null
        
    }
	
	$data+=New-Object psobject -property @{Alias=$alias;EventId=$eventId;Event=$event;'Session ID'=$sessionID;'Source IP'=$sourceIP;Time=$time}

}

[Environment]::SetEnvironmentVariable("RDLogsCollected","$date","Machine")

$data | select Alias,EventID,Event,'Session ID','Source IP',Time | Sort Time | export-csv C:\RDConnectionReports\Reports\$filename -notypeinformation