################################################################################
##      This script determines if ther are jobs in a print queue and controls
## a smart outlet accordingly -- turning the outlet on if there are jobs
## in the queue, turning it off if there are no jobs in the queue.
## Editable Variables:
# Name of the print queue
$strPrintQueueName = "WF-7520 Series(Network)"
# OpenHAB2 REST address for the outlet to be controlled
$strOH2OutletURI = "http://oh2host.domain.ccTLD:8080/rest/items/Outlet1"
# Number of seconds to pause between queue length checks before powering off outlet
$iPauseDuraction = 600
################################################################################
# Get number of jobs in print queue
$iItemsInQueue = get-printjob -PrinterName $strPrintQueueName | measure

# Get current state of outlet
try{
	$strOutletState = Invoke-WebRequest -URI '$strOH2OutletURI/state' -ContentType "text/plain" -Method GET 
}
catch{
	Write-Host "Status Code:" $_.Exception.Response.StatusCode.value__ 
	Write-Host "Status Description:" $_.Exception.Response.StatusDescription
	Write-Host "Status Message:" $_.Exception.Message
	Write-Host "Status Response:" $_.Exception.Response
	Write-Host "Status Status:" $_.Exception.Status
}


if($iItemsInQueue.Count -eq 0 -And $strOutletState.Content-eq "ON"){			# No items are in print queue, but outlet is on. Turn it off after a pause
	Write-Host "No items in the print queue but outlet is on; the outlet will be turned off."
	try{
		Start-Sleep $iPauseDuration
		$iItemsInQueue2 = get-printjob -PrinterName $strPrintQueueName | measure		# Check queue again to ensure we are not turning outlet off while someone is submitting print job
		if($iItemsInQueue2.Count -eq 0){
			Write-Host "Still no items in queue, turning off outlet."
			Invoke-WebRequest -URI $strOH2OutletURI -ContentType "text/plain" -Method POST -Body 'OFF'
		}
	}
	catch{
		Write-Host "Status Code:" $_.Exception.Response.StatusCode.value__ 
		Write-Host "Status Description:" $_.Exception.Response.StatusDescription
		Write-Host "Status Message:" $_.Exception.Message
		Write-Host "Status Response:" $_.Exception.Response
		Write-Host "Status Status:" $_.Exception.Status
	}
}
elseif ($iItemsInQueue.Count -gt 0 -And $strOutletState.Content-eq "OFF"){		# Items are in print queue and outlet is not presently on
	Write-Host "Items are in the print queue but outlet is off; the outlet will be turned on."
	try{
		Invoke-WebRequest -URI $strOH2OutletURI' -ContentType "text/plain" -Method POST -Body 'ON'
	}
	catch{
		Write-Host "Status Code:" $_.Exception.Response.StatusCode.value__ 
		Write-Host "Status Description:" $_.Exception.Response.StatusDescription
		Write-Host "Status Message:" $_.Exception.Message
		Write-Host "Status Response:" $_.Exception.Response
		Write-Host "Status Status:" $_.Exception.Status
	}
}
else{
	Write-Host $iItemsInQueue.Count " items in the print queue. The outlet is currently " $strOutletState.Content
}
