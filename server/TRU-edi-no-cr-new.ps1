. .\TRU-config.ps1
. ..\prog\common\CSAccess.ps1

################################
# Set non-stopping error to STOP
################################

$ErrorActionPreference = "Stop"

###################
# Trapping function
###################


trap {
	$smtp = new-object net.mail.smtpclient($Script:MailRelay)
   	$smtp.send($Script:MailSender, $Script:MailRecipient, "TRU EDI execution error", $($error[0].tostring() + $error[0].invocationinfo.positionmessage))
	break
}


#################
# Common function
#################

function GenFilename {
#	"geoffrey390fix_rl294_MCP.NA.MC00DEFV.PXSFIL." + (Get-Date -format "DyyMMdd.THHmmss")
	"geoffrey390fix_rl338_MCP.NA.MC00DEFV.PXSFIL." + (Get-Date -format "DyyMMdd.THHmmss")
}

function Check-Dropoff {
	param ($CheckFile)	
	
	$valReturn = 0
	
	trap [System.Management.Automation.MethodInvocationException] {
		switch -regex ($error[0].toString()) {
			# Server not responsing
			"Unable to connect to the remote server" {set-variable -name valReturn -scope 1 -value -1; }
			# Credential incorrect
			"(530)" {set-variable -name valReturn -scope 1 -value -2; }
			# File not find
			"(450)" {set-variable -name valReturn -scope 1 -value -3; }
		}
		continue
	}

	$ftpUri = "ftp://10.10.1.27/verification/"
	$ftpUser = "truivx"
	$ftpPass = "gn>~_{y;?w~?"

	$ftpUri = $ftpUri + $CheckFile
	$ftp = [system.net.ftpwebrequest]::create($ftpUri)
	$ftp.Credentials.username = $ftpUser
	$ftp.Credentials.password = $ftpPass

	$ftp.method = "LIST"
	$ftpResponse = $ftp.getresponse()
	$ftpResponse.close()
#	out-host -InputObject $ftpUri
	if ($valReturn -eq 0) {
		# $ftpResponse.close()
		$ftp = $null
		if ($CheckFile -ne $null) {
			SendAlert "$CheckFile is in DropOff"
		}
	}
	#Out-Host -InputObject $valReturn
	switch ($valReturn) {
		-1 {SendAlert "DropOff not responsing"}
		-3 {SendAlert "$CheckFile not found in DropOff"}
		-2 {SendAlert "Credential to logon DropOff is incorrect"}
	}

 #   $valReturn = 0
	$valReturn
}

####################
# Generate Feed Part
####################

function PushToSftp {
<##########
###########

	$QueryCAML = 
@"
	<Where>
		<And>
      		<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
      		<Eq><FieldRef Name='FileType' /><Value Type='Text'>S</Value></Eq>
   		</And>
	</Where>
"@
	$Query = New-Object Microsoft.SharePoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:FeedAckTable.GetItems($Query)
	
###########
##########>

	$QueryCAML = 
@"
    <View>
    <Query>
	<Where>
		<And>
      		<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
      		<Eq><FieldRef Name='FileType' /><Value Type='Text'>S</Value></Eq>
   		</And>
	</Where>
    </Query>
    </View>
"@
	$Result = GetCAMLResult $Script:FeedAckTable $QueryCAML

#	Copy-Item -Path (Join-Path $Script:NotVert $Result[0]["Filename"]) -destination (Join-Path $Script:Vert $Result[0]["Filename"])
	
	"####"
	"PUSHING TO SFTP..."
	"###"


	.\TRU\sftp\sftpProg\upload_verification.bat
	
	$ID = @($Result.FieldValues)[0].ID
	$Filename = @($Result.FieldValues)[0].Filename
	$SPItem = $Script:FeedAckTable.GetItemById($ID)
	$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}",[datetime]::utcnow)
	$SPItem.Update()
    $Script:ctx.executeQuery()

#	[void] $(Check-DropOff $SPItem["Filename"])
	[void] $(Check-DropOff $Filename)
	
	$Script:RunningState = 3

<##########
###########
	$SPItem = $Script:StateTable.items[0]
	$SPItem["State"] = $Script:RunningState
	$SPItem["CurrentFSN"] = ++$Script:CurrentFSN
	$SPItem.Update()
###########
##########>

	$SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
	$SPItem["State"] = $Script:RunningState
	$SPItem["CurrentFSN"] = ++$Script:CurrentFSN
	$SPItem.Update()
    $Script:ctx.executeQuery()
	$Query = $null
}

function GenerateFeed {
	$FeedFileName = GenFilename
<##########
###########
	$QueryCAML = 
@"
	<Where>
		<And>
      		<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
      		<Eq><FieldRef Name='FileType' /><Value Type='Text'>S</Value></Eq>
   		</And>
	</Where>
"@
###########
##########>
	$QueryCAML = 
@"
    <View>
    <Query>
	<Where>
		<And>
      		<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
      		<Eq><FieldRef Name='FileType' /><Value Type='Text'>S</Value></Eq>
   		</And>
	</Where>
    </Query>
    </View>
"@
<##########
###########
	$Query = New-Object Microsoft.SharePoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:FeedAckTable.GetItems($Query)
###########
##########>

    $Result = GetCAMLResult $Script:FeedAckTable $QueryCAML
	if ($Result.count -eq 0) {
<##########
###########
		$SPItem = $Script:FeedAckTable.Items.Add()
		$SPItem["FSN"] = $Script:CurrentFSN
		$SPItem["FileType"] = "S"
		$SPItem["FSNReusedNum"] = 1
		$SPItem["Filename"] = $FeedFileName
		$SPItem["Title"] = $FeedFileName
###########
##########>

        $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $SPItem = $Script:FeedAckTable.AddItem($listItemInfo)
		$SPItem["FSN"] = $Script:CurrentFSN
		$SPItem["FileType"] = "S"
		$SPItem["FSNReusedNum"] = 1
		$SPItem["Filename"] = $FeedFileName
		$SPItem["Title"] = $FeedFileName
        $SPItem.Update()
        $Script:ctx.load($Script:FeedAckTable)
        $Script:ctx.executeQuery()
    	$listItemInfo = $null
	}
	else {
        $ID = (@($Result.FieldValues)[0]).ID
		$SPItem = $Script:FeedAckTable.GetItemById($ID)
		$SPItem["Filename"] = $FeedFileName
		$SPItem["Title"] = $FeedFileName
        $SPItem.Update()
        $Script:ctx.executeQuery()
        
        # Check if file attachment exists?
	}
<##########
###########
	$SPItem.Update()
###########
##########>

	$SPItem1 = $SPItem		
	$SPItem = $null		

<##########
###########
	$Query = $null
	$QueryCAML = 
@"
	<Where>
		<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
    </Where>
	<OrderBy>
   		<FieldRef Name='SKN' Ascending='True' />
	</OrderBy>

"@
	$Query = New-Object Microsoft.SharePoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:RecordSentTable.GetItems($Query)
	$Query = $null
###########
##########>

	$QueryCAML = 
@"
    <View>
    <Query>
	<Where>
		<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
    </Where>
	<OrderBy>
   		<FieldRef Name='SKN' Ascending='True' />
	</OrderBy>
    </Query>
    </View>

"@
    $Result = GetCAMLResult $Script:RecordSentTable $QueryCAML

	OutFeed $Result $Script:CurrentFSN (Join-Path $Script:TempDir $FeedFileName)

<##########
###########

	$Content = Get-Content -encoding Byte -path (Join-Path $Script:NotVert $FeedFileName)
	$SPItem1.attachments.add($FeedFileName, $Content)
	$SPItem1.Update()	
###########
##########>
    
    # Using web service to add file into FeedActTable, CSOM doesn't support

    $fStream=[System.IO.File]::OpenRead($(Join-Path $Script:NotVert $FeedFileName))
    [String] $fName=$fStream.Name
    [System.Byte[]]$bytes=New-Object -TypeName System.Byte[] -ArgumentList $fStream.Length
    [void] $fStream.Read($bytes, 0, [int]$fStream.Length)
    $fStream.Close()
    $listname = "FeedAckTable"

    # Get back the record ID
    $Script:ctx.load($SPItem1)
    $Script:ctx.executeQuery()
    $ID = @($SPItem1.FieldValues)[0].ID
    $listWebServiceReference = New-WebServiceProxy -Uri $Script:WSuri -UseDefaultCredential
    [void] $listWebServiceReference.AddAttachment($listName,$ID,$fName,$bytes)

<##########
###########
	$Script:RunningState = 2
	$SPItem = $Script:StateTable.items[0]
	$SPItem["State"] = $Script:RunningState
	$SPItem.Update()
###########
##########>
	$Script:RunningState = 2
    $SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
	$SPItem["State"] = $Script:RunningState
	$SPItem.Update()
    $Script:ctx.executeQuery()
}

function OutFeed ($Result1, $CurrentFSN, $FeedFileName) {

#	exit



	$RecordTotal = 0
	$FeedHeader = "HEADER" + (Get-Date -format "ddMMyyyy") + ("{0:000000}" -f $CurrentFSN) 
	#$HeaderLength = 294 - $FeedHeader.length
	#+ " ".padright(276)  #murphy: 17-APR-2009
#	$FeedHeader = $FeedHeader.padright(294)
	$FeedHeader = $FeedHeader.padright(294) + "`n"
	
#	Out-File -filepath $FeedFileName -inputobject $FeedHeader -encoding ascii -append
	Add-Content $FeedFileName -value $([Byte[]]([Char[]]$FeedHeader)) -encoding byte

<##########
###########

	$Result1 | foreach {
		$SPItem = $_
		$SPItem['Date'] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
		$SPItem.Update()
		
		$QueryCAML =
@"
			<Where>
  				<Eq><FieldRef Name='SKN_N' /><Value Type='Number'>$($_["SKN"])</Value></Eq>
			</Where>
			<OrderBy>
   				<FieldRef Name='ID' Ascending='False' />
			</OrderBy>
"@
		$Query = new-object Microsoft.SharePoint.SPQuery
		$Query.query = $QueryCAML
		$Result = $Script:TruSPList.GetItems($query)
###########
##########>

	foreach ($r in $Result1.FieldValues) {
        $SPItem = $Script:RecordSentTable.GetItemById($r.ID)
		$SPItem['Date'] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
        $SPItem.Update()
        $Script:ctx.executeQuery()
		
		$QueryCAML =
@"
            <View>
            <Query>
			<Where>
  				<Eq><FieldRef Name='SKN_N' /><Value Type='Number'>$($r["SKN"])</Value></Eq>
			</Where>
			<OrderBy>
   				<FieldRef Name='ID' Ascending='False' />
			</OrderBy>
            </Query>
            </View>
"@
		$Result = GetCAMLResult $Script:TruSPList $QueryCAML
        $Result = @($Result.FieldValues)

#		$TT = -1
	
#		$Result | foreach {$TT++}
		
#		$TT

<##########
###########
# No used 20171119
    	$ChkOpt = 0
		"Item: " + $_["SKN"]
		$Result | foreach {
			
			$_["ID"]
			if ($ChkOpt -eq 0) {$_["OptmzAble"]}
			$_["VrDate"]
			$ChkOpt++
			"Yello"
		}
		"`n"
###########
##########>

		$RecordTotal += $Result[0]["SKN_N"]

		if ($Result[0]["OptmzAble"] -eq $null) {
			$OptOpt = "20 - "
		}
		else {
			$OptOpt = $Result[0]["OptmzAble"]
		}
	
		
		$FeedLine = `
		("{0:000000000000000}" -f $Result[0]["SKN_N"]).substring(0,15) + `
#		("{0:D15}" -f [int64]$Result[0]["UPC"]).substring(0,15) + `
		(("{0,15}" -f ("{0:000000000000000}" -f $Result[0]["UPC"])).substring(0,15)).replace(" ", "0") + `
		("{0,-8}" -f $Result[0]["SKU"]).substring(0,8) + `
#		("{0,-120}" -f $Result[0]["Descr"]).substring(0,120) + `
		("{0,-120}" -f $Result[0]["Title"]).substring(0,120) + `
#		("{0,6}" -f (("{0:0000.00}" -f $Result[0]["CsLngt"]) -replace "\.")).substring(0,6) + `
#		("{0,6}" -f (("{0:0000.00}" -f $Result[0]["CsWdth"]) -replace "\.")).substring(0,6) + `
		("{0,6}" -f (("{0:0000.00}" -f $Result[0]["CsWdth"]) -replace "\.")).substring(0,6) + `
		("{0,6}" -f (("{0:0000.00}" -f $Result[0]["CsLngt"]) -replace "\.")).substring(0,6) + `
		("{0,6}" -f (("{0:0000.00}" -f $Result[0]["CsHt"]) -replace "\.")).substring(0,6) + `
#		("{0,5}" -f (("{0:0000.0}" -f $Result[0]["CsVol"]) -replace "\.")).substring(0,5) + `
#		("{0,5}" -f (("{0:0000.0}" -f [math]::Round($Result[0]["CsVol_N"].split("#")[1],1)) -replace "\.")).substring(0,5) + `
#       $(if ($Result[0]["CsVol_N"].split("#")[1] -ne "") {("{0,5}" -f (("{0:0000.0}" -f [double]$Result[0]["CsVol_N"].split("#")[1]) -replace "\.")).substring(0,5)} else {"     "}) + `
#		("{0,5}" -f (("{0:0000.0}" -f [double]$Result[0]["CsVol_N"]) -replace "\.")).substring(0,5) + `
        $(if ($Result[0]["CsVol_N"] -ne "") {("{0,5}" -f (("{0:0000.0}" -f [double]$Result[0]["CsVol_N"]) -replace "\.")).substring(0,5)} else {"     "}) + `
		(("{0,2}" -f $Result[0]["CsMsr"]).toupper()).substring(0,2) + `
		("{0,5}" -f (("{0:000.00}" -f $Result[0]["CsWt"]) -replace "\.")).substring(0,5) + `
		(("{0,2}" -f $Result[0]["CsWtMsr"]).toupper()).substring(0,2) + `
		("{0,4}" -f (("{0:00.00}" -f $Result[0]["UnWdth"]) -replace "\.")).substring(0,4) + `
		("{0,4}" -f (("{0:00.00}" -f $Result[0]["UnLngt"]) -replace "\.")).substring(0,4) + `
#		("{0,4}" -f (("{0:00.00}" -f $Result[0]["UnLngt"]) -replace "\.")).substring(0,4) + `
#		("{0,4}" -f (("{0:00.00}" -f $Result[0]["UnWdth"]) -replace "\.")).substring(0,4) + `
		("{0,4}" -f (("{0:00.00}" -f $Result[0]["UnHt"]) -replace "\.")).substring(0,4) + `
		(("{0,2}" -f $Result[0]["UnMsr"]).toupper()).substring(0,2) + `
#		("{0,5}" -f (("{0:0000.0}" -f $Result[0]["UnVol"]) -replace "\.")).substring(0,5) + `
#		("{0,5}" -f (("{0:0000.0}" -f [math]::Round($Result[0]["UnVol_N"].split("#")[1],1)) -replace "\.")).substring(0,5) + `
#       $(if ($Result[0]["UnVol_N"].split("#")[1] -ne "") {("{0,5}" -f (("{0:0000.0}" -f [double]$Result[0]["UnVol_N"].split("#")[1]) -replace "\.")).substring(0,5)} else {"     "}) + `
#        ("{0,5}" -f (("{0:0000.0}" -f [double]$Result[0]["UnVol_N"]) -replace "\.")).substring(0,5) + `
        $(if ($Result[0]["UnVol_N"] -ne "") {("{0,5}" -f (("{0:0000.0}" -f [double]$Result[0]["UnVol_N"]) -replace "\.")).substring(0,5)} else {"     "}) + `
		("{0,5}" -f (("{0:000.00}" -f $Result[0]["UnWt"]) -replace "\.")).substring(0,5) + `
		(("{0,2}" -f $Result[0]["UnWtMsr"]).toupper()).substring(0,2) + `
		("{0,7}" -f ("{0:0000000}" -f $Result[0]["MPkQty"])).substring(0,7) + `
		("{0,7}" -f ("{0:0000000}" -f $Result[0]["IPQty"])).substring(0,7) + `
		("{0,1}" -f $(if ($Result[0]["Exmpt"] -eq "True"){"Y"} elseif ($Result[0]["Exmpt"] -eq "False"){"N"})) + `
		("{0,1}" -f $(if ($Result[0]["BtyReq"] -eq "True"){"Y"} elseif ($Result[0]["BtyReq"] -eq "False"){"N"})) + `
		("{0,1}" -f $(if ($Result[0]["BtyIncl"] -eq  "True"){"Y"} elseif ($Result[0]["BtyIncl"] -eq "False"){"N"})) + `
		("{0,-9}" -f $Result[0]["BtSz1"]).substring(0,9) + `
		("{0,2}" -f ("{0:00}" -f $Result[0]["BtQty1"])).substring(0,2) + `
		("{0,-9}" -f $Result[0]["BtSz2"]).substring(0,9) + `
		("{0,2}" -f ("{0:00}" -f $Result[0]["BtQty2"])).substring(0,2) + `
		("{0,-9}" -f $Result[0]["BtSz3"]).substring(0,9) + `
		("{0,2}" -f ("{0:00}" -f $Result[0]["BtQty3"])).substring(0,2) + `
		("{0,14}" -f ("{0:MMddyyyyHHmmss}" -f $Result[0]["RcDate"])).substring(0,14) + `
		("{0,14}" -f ("{0:MMddyyyyHHmmss}" -f $Result[0]["VrDate"])).substring(0,14) + `
		$OptOpt.substring(0,2) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {$Result[0]["PackMaterialType"].substring(0,2)} else {" " * 2}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,6}" -f (("{0:0000.00}" -f $Result[0]["OptCaseLen"]) -replace "\.")).substring(0,6)} else {" " * 6}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,6}" -f (("{0:0000.00}" -f $Result[0]["OptCaseWdth"]) -replace "\.")).substring(0,6)} else {" " * 6}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,6}" -f (("{0:0000.00}" -f $Result[0]["OptCaseHght"]) -replace "\.")).substring(0,6)} else {" " * 6}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,5}" -f (("{0:000.00}" -f $Result[0]["OptCaseWght"]) -replace "\.")).substring(0,5)} else {" " * 5}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,4}" -f (("{0:00.00}" -f $Result[0]["OptItemLen"]) -replace "\.")).substring(0,4)} else {" " * 4}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,4}" -f (("{0:00.00}" -f $Result[0]["OptItemWdth"]) -replace "\.")).substring(0,4)} else {" " * 4}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,4}" -f (("{0:00.00}" -f $Result[0]["OptItemHght"]) -replace "\.")).substring(0,4)} else {" " * 4}) + `
		$(if ($OptOpt.substring(0,2) -eq "10") {("{0,5}" -f (("{0:000.00}" -f $Result[0]["OptItemWght"]) -replace "\.")).substring(0,5)} else {" " * 5})
			
		$FeedLine += "`n"
		
#		out-file $FeedFileName -inputobject $FeedLine -encoding ascii -append
		Add-Content $FeedFileName -value $([Byte[]]([Char[]]$FeedLine)) -encoding byte

<##########
###########

		$SPItem = $Result[0]
		$SPItem["FeedBack"] = $Script:CodeDesc["1000"]
		$SPItem["SendToTRU"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
#		$SPItem["AckFromTRU"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::MaxValue)
		$SPItem["AckFromTRU"] = "9999-12-31T00:00:00Z"
		$SPItem.Update()
		$Query = $null
###########
##########>

		$SPItem = $Script:TruSPList.GetItemById($Result[0].ID)
		$SPItem["FeedBack"] = $Script:CodeDesc["1000"]
		$SPItem["SendToTRU"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
#		$SPItem["AckFromTRU"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::MaxValue)
		$SPItem["AckFromTRU"] = "9999-12-31T00:00:00Z"
	    $SPItem.Update()
        $Script:ctx.executeQuery()
	}
	
#	$FeedTrailer = "TRAILER" + ("{0:0000000000}" -f ($Result1.count+2)) + ("{0:0000000000}" -f $RecordTotal) + " ".padright(267)
	$FeedTrailer = "TRAILER" + ("{0:0000000000}" -f ($Result1.count+2)) + ("{0:0000000000}" -f $RecordTotal) + " ".padright(267) + "`n"
#	Out-File $FeedFileName -inputobject $FeedTrailer -encoding ascii -append
	
	Add-Content $FeedFileName -value $([Byte[]]([Char[]]$FeedTrailer)) -encoding byte
	
	Remove-Item -path $($Script:NotVert + "\*")
	Move-Item -path $FeedFileName -destination (Join-Path $Script:NotVert (Split-Path -leaf $FeedFileName)) -force -erroraction SilentlyContinue
	Remove-Item -path $($Script:TempDir + "\*") -exclude verify_ack

#	exit
}

#function FindRecordToSend {
#	$QueryCAML = `
#@"
#			<Where>
#  				<Eq><FieldRef Name='Processed' /><Value Type='Text'>N</Value></Eq>
#				<IsNotNull><FieldRef Name='VrDate' /></IsNotNull>
#			</Where>
#			<OrderBy>
#   				<FieldRef Name='SKN_N' Ascending='True' />
#			</OrderBy>
#"@
#	$Query = new-object Microsoft.SharePoint.SPQuery
#	$Query.query = $QueryCAML
#	$Result = $Script:TruSPList.GetItems($query)
#	$Result | foreach {
#		$QueryCAML = 
#@"
#			<Where>
#				<And>
#      				<Eq><FieldRef Name='SKN' /><Value Type='Text'>$($_["SKN_N"])</Value></Eq>
#      				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
#   				</And>
#			</Where>
#"@
#		$Query1 = new-object Microsoft.SharePoint.SPQuery
#		$Query1.query = $QueryCAML
#		$Result1 = $Script:RecordSentTable.GetItems($Query1)
#		if ($Result1.count -eq 0) {
#			$SPItem = $Script:RecordSentTable.Items.Add()
#			$SPItem["SKN"] = $_["SKN_N"]
#			$SPItem["FSN"] = $Script:CurrentFSN
#			$SPItem["FSNReusedNum"] = 1
#			$SPItem.Update()
#			$SPItem = $null
#		}
#		$SPItem = $_
#		$SPItem["Processed"] = "Y"
#		$SPItem.Update()
#		$Query1 = $null
#	}
#	$Query = $null
	
#	$QueryCAML = 
#@"
#	<Where>
#		<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
#	</Where>
#"@
#	$Query = New-Object Microsoft.Sharepoint.SPQuery
#	$Query.query = $QueryCAML
#	$Result = $Script:RecordSentTable.GetItems($Query)

#	if ($Result.count -ne 0) {
#		$Script:RunningState = 1
#		$SPItem = $Script:StateTable.items[0]
#		$SPItem["State"] = $Script:RunningState
#		$SPItem.Update()
#		$SPItem = $null
#	}
#	else {
#		$Script:RunningState = 3
#		$SPItem = $Script:StateTable.items[0]
#		$SPItem["State"] = $Script:RunningState
#		$SPItem.Update()
#		$SPItem = $null
#	}
#	$Query = $null
#}

function FindRecordToSend {
	$CurrentTime = [datetime]::Now
	$Condition1 = $CurrentTime.hour -ge 20
#	$Condition1 = $CurrentTime.hour -ge 16
	$Condition2 = (($CurrentTime.dayofweek -ne "Saturday") -and ($CurrentTime.dayofweek -ne "Sunday"))

<##########
	$QueryCAML = 
@"
	<Where>
		<And>
			<And>
				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$($Script:CurrentFSN - 1)</Value></Eq>			
				<Eq><FieldRef Name='FileType' /><Value Type='Text'>S</Value></Eq>	
			</And>
			<Eq><FieldRef Name='FSNReusedNum' /><Value Type='Number'>1</Value></Eq>
		</And>
	</Where>
"@
    "Hello"
##########>

	$QueryCAML = 
@"
    <View>
        <Query>
	        <Where>
		        <And>
			        <And>
				        <Eq><FieldRef Name='FSN' /><Value Type='Number'>$($Script:CurrentFSN - 1)</Value></Eq>			
				        <Eq><FieldRef Name='FileType' /><Value Type='Text'>S</Value></Eq>	
			        </And>
			        <Eq><FieldRef Name='FSNReusedNum' /><Value Type='Number'>1</Value></Eq>
		        </And>
	        </Where>
        </Query>
        <RowLimit>1</RowLimit> 
    </View>
"@

<##########
	$Query = new-object Microsoft.SharePoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:FeedAckTable.GetItems($query)
	$Query = $null
	$Condition3 = ($Result[0]['Date'].date -lt $CurrentTime.date)
##########>

	$Result = GetCAMLResult $Script:FeedAckTable $QueryCAML 
	$Condition3 = (@($Result.FieldValues)[0]["Date"].date -lt $CurrentTime.date)

	if (($Condition1 -and $Condition2) -and $Condition3) {
#	if ($true) {	
		[void] $(Check-DropOff @($Result.FieldValues)[0]["Filename"])

<##########
		$QueryCAML = `
@"
			<Where>
  				<Eq><FieldRef Name='Processed' /><Value Type='Text'>N</Value></Eq>
				<IsNotNull><FieldRef Name='VrDate' /></IsNotNull>
			</Where>
			<OrderBy>
   				<FieldRef Name='SKN_N' Ascending='True' />
			</OrderBy>
"@
##########>

		$QueryCAML =
@"
    <View>
        <Query>
			<Where>
  				<Eq><FieldRef Name='Processed' /><Value Type='Text'>N</Value></Eq>
				<IsNotNull><FieldRef Name='VrDate' /></IsNotNull>
			</Where>
			<OrderBy>
   				<FieldRef Name='SKN_N' Ascending='True' />
			</OrderBy>
        </Query>
    </View>
"@

<##########
		$Query = new-object Microsoft.SharePoint.SPQuery
		$Query.query = $QueryCAML
		$Result = $Script:TruSPList.GetItems($query)
##########>

        $Result = GetCAMLResult $Script:TruSPList $QueryCAML

<##########

		$Result | foreach {

			$QueryCAML = 
@"
			<Where>
				<And>
      				<Eq><FieldRef Name='SKN' /><Value Type='Text'>$($_["SKN_N"])</Value></Eq>
      				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
   				</And>
			</Where>
"@
##########>

		foreach ($r in $Result.FieldValues) {
			$QueryCAML = 
@"
    <View>
        <Query>
			<Where>
				<And>
      				<Eq><FieldRef Name='SKN' /><Value Type='Text'>$($r["SKN_N"])</Value></Eq>
      				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentFSN</Value></Eq>
   				</And>
			</Where>
        </Query>
    </View>
"@

<##########
			$Query1 = new-object Microsoft.SharePoint.SPQuery
			$Query1.query = $QueryCAML
			$Result1 = $Script:RecordSentTable.GetItems($Query1)
##########>
            
            $Result1 = GetCAMLResult $Script:RecordSentTable $QueryCAML

			if ($Result1.count -eq 0) {
<##########
				$SPItem = $Script:RecordSentTable.Items.Add()
				$SPItem["SKN"] = $_["SKN_N"]
				$SPItem["FSN"] = $Script:CurrentFSN
				$SPItem["FSNReusedNum"] = 1
				$SPItem.Update()
				$SPItem = $null
##########>
#                $list = $Script:Ctx.web.Lists.GetByTitle("RecordSentTable")
                $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                $SPItem = $Script:RecordSentTable.AddItem($listItemInfo)
                $SPItem["SKN"] = $r["SKN_N"]
                $SPItem["FSN"] = $Script:CurrentFSN
			    $SPItem["FSNReusedNum"] = 1
                $SPItem.Update()
                $Script:ctx.load($Script:RecordSentTable)
                $Script:ctx.executeQuery()
			    $listItemInfo = $null
			}
			$SPItem = $Script:TruSPList.GetItemById($r.ID)
			$SPItem["Processed"] = "Y"
			$SPItem.Update()
            $Script:ctx.executeQuery()
			$Query1 = $null
		}
		$Query = $null

		$Script:RunningState = 1
<##########
###########

		$SPItem = $Script:StateTable.items[0]
###########
##########>

		$SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
		$SPItem["State"] = $Script:RunningState
		$SPItem.Update()
        $Script:ctx.executeQuery()
		$SPItem = $null
	}
	else {
		$Script:RunningState = 3
<##########
###########
		$SPItem = $Script:StateTable.items[0]
###########
##########>
		$SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
		$SPItem["State"] = $Script:RunningState
		$SPItem.Update()
        $Script:ctx.executeQuery()
		$SPItem = $null
	}
	$Query = $null
}

#######################
# Receive Ack part
#######################

function SortAck {
	$Fileinfo = Get-ChildItem -Path $Script:NotAck | foreach {
		$h = @{}
		$Header = Get-Content -TotalCount 1 (join-path $Script:NotAck $_)

# 7-24-2013 change
# TRU changes the header and trailer without following an exact 261 length
#
#		if ($Header.length -eq 261) {
#		   $Field1 = $Header.substring(0,6)
#		   $Field2 = $Header.substring(14,6)
#		}
#		else {
#		   $h.fsn = -1
#		}

		$Field1 = $Header.substring(0,6)
		$Field2 = $Header.substring(14,6)
		
		if (($Field1 -match "HEADER") -and ($Field2 -match "\d{6}")) {
			$h.fsn = [int]$Field2
		}
		else {
			$h.fsn = -1
		}
		$h.filename = $_.name
		$h
	}
	#if ($err) {-1} else {, @($Fileinfo | sort-object -Property @{expression={$_.fsn}})}
	, @($Fileinfo | Sort-Object -Property @{expression={$_.fsn}})
}

function ValidateAck ($Filename) {
	#
	# Checking the content
	#

	$RecordTotalThis = 0
	$AllLine = Get-Content (Join-Path $Script:NotAck $Filename) -encoding string
	
	# This is ugly as the ack from TRU has a strange line of char after TRAILER line
	if ($AllLine[-1] -match "^TRAILER") {
		$Adjust = 1
	}
	elseif ($AllLine[-2] -match "^TRAILER") {
		$Adjust = 2
	}
	else {
		return -4
	}
	
	$AllLine[1 .. $($AllLine.length - (1 + $Adjust))] | foreach {
		$Field1 = $Field2 = $null

# 7-24-2013
# TRU changes without following a 261 record len

#		if ($_.length -eq 261) {
#		   $Field1 = $_.substring(0,15)
#		   $Field2 = $_.substring(158,3)
#		}

		$Field1 = $_.substring(0,15)
		$Field2 = $_.substring(158,3)

		if (($Field1 -match "\d{15}") -and ($Field2 -match "\d{3}")) {
			if ("100", "101", "102", "103" -contains $Field2) {
				if ($AllLine.length -eq 3) {
					$RecordTotalThis = 0
				}
				else {
					$err = $true
				}
			}
			else {
				$RecordTotalThis += [double] $Field1
			}
		}
		else {
			# Content error
			$err = $true
		}
	}
	if ($err) {return -3}

	#
	# Checking the trailer
	#
	$Field1 = $Field2 = $null
# 7-24-2013
# TRU changes the record len without following the 261 specification


#	if ($AllLine[-$($Adjust)].length -eq 261) {
#		$Field1 = $AllLine[-$($Adjust)].substring(7,10)
#		$Field2 = $AllLine[-$($Adjust)].substring(17,10)
#	}

	$Field1 = $AllLine[-$($Adjust)].substring(7,10)
	$Field2 = $AllLine[-$($Adjust)].substring(17,10)

	if (($Field1 -match "\d{10}") -and ($Field2 -match "\d{10}")) {
		$Line = [double]$Field1 - 2
		$RecordTotal = [double]$Field2

		if ($RecordTotal -eq 9999999999) {
			$RecordTotal = $RecordTotalThis
		}
		if (($Line -eq $($AllLine.length - (1 + $Adjust))) -and ($RecordTotal -eq $RecordTotalThis)) {

			return 0
		}
		else {
			# Trailer error
			return -4
		}
	}
	else {
		# Trailer error
		return -4
	}
}

function PullFromSftp {
	"###"
	"# PULL FROM SFTP"
	"###"
	#"1 "+ [datetime]::Now.touniversaltime() | Out-File $Logfile -encoding ascii -append
	.\TRU\sftp\sftpProg\download_acks.bat
	#"2 "+ [datetime]::Now.touniversaltime() | Out-File $Logfile -encoding ascii -append
	if ((Get-ChildItem $Script:NotAck) -eq $null) {
		$Script:RunningState = 6
	}
	else {
<##########
###########

		$QueryCAML = 
@"
			<Where>
				<And>
  					<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
					<Eq><FieldRef Name='FileType' /><Value Type='Text'>R</Value></Eq>
				</And>
			</Where>
"@
		$Query = New-Object Microsoft.Sharepoint.SPQuery
		$Query.query = $QueryCAML
		$Result = $Script:FeedAckTable.GetItems($Query)
		$Query = $null
###########
##########>

		$QueryCAML = 
@"
            <View>
            <Query>
			<Where>
				<And>
  					<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
					<Eq><FieldRef Name='FileType' /><Value Type='Text'>R</Value></Eq>
				</And>
			</Where>
            </Query>
            </View>
"@
		$Result = GetCAMLResult $Script:FeedAckTable $QueryCAML

		if ($Result.count -ne 0) {
<##########
###########
			$SPItem = $Result[0]
			$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
			$SPItem.Update()
###########
##########>

			$ID = @($Result.FieldValues)[0].ID
            $SPItem = $Script:FeedAckTable.GetItemById($ID)
			$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
			$SPItem.Update()
            $Script:ctx.executeQuery()

		    # Delete the attachment by Web service first
            $listName = "FeedAckTable"
            $listWebServiceReference = New-WebServiceProxy -Uri $WSuri -UseDefaultCredential
            [System.Xml.XmlNode]$xmlNode=$listWebServiceReference.GetAttachmentCollection($listName,$ID) 
            foreach($node in $xmlNode.Attachment) {
                #$node
                [void] $listWebServiceReference.DeleteAttachment($listName,$ID,$node) 
            }

            # Add the attachment by Web service
			$Result = SortAck
			$Script:CurrentACKFile = $Result[0].filename
            $listName = "FeedAckTable"

            $fStream=[System.IO.File]::OpenRead($(Join-Path $Script:NotAck $Script:CurrentACKFile))
            [String] $fName=$fStream.Name
            [System.Byte[]]$bytes=New-Object -TypeName System.Byte[] -ArgumentList $fStream.Length
            [void] $fStream.Read($bytes, 0, [int]$fStream.Length)
            $fStream.Close()
            $listWebServiceReference = New-WebServiceProxy -Uri $WSuri -UseDefaultCredential
            [void] $listWebServiceReference.AddAttachment($listName,$ID,$fName,$bytes)

<##########
###########
			$Script:RunningState = 4
			$SPItem = $StateTable.items[0]
			$SPItem["State"] = $Script:RunningState
			$SPItem.Update()
###########
##########>

			$Script:RunningState = 4
			$SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
			$SPItem["State"] = $Script:RunningState
			$SPItem.Update()
            $Script:ctx.executeQuery()
		}
		else {
			$Result = SortAck
			$Script:CurrentACKFile = $Result[0].filename
			if ($Result[0].fsn -lt 0) {
				# Ack Header Error
				$Script:RunningState = -2

<##########
###########
				$SPItem = $StateTable.items[0]
				$SPItem["State"] = $Script:RunningState
				$SPItem.Update()		
###########
##########>
                $SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
				$SPItem["State"] = $Script:RunningState
				$SPItem.Update()		
                $Script:ctx.executeQuery()
			}
			else {
				$Result1 = ValidateAck $Result[0].filename
				if ($Result1 -ne 0) {
					# Ack trailer or content error
					$Script:RunningState = $Result1

<##########
###########
					$SPItem = $StateTable.items[0]
					$SPItem["State"] = $Script:RunningState
					$SPItem.Update()		
###########
##########>
                    $SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
				    $SPItem["State"] = $Script:RunningState
				    $SPItem.Update()		
                    $Script:ctx.executeQuery()
				}
				else {
					if ($Result[0].fsn -ne $Script:CurrentACK) {
						# Ack fsn out of sync
						$Script:RunningState = -1

<##########
###########
						$SPItem = $StateTable.items[0]
						$SPItem["State"] = $Script:RunningState
						$SPItem.Update()
###########
##########>
                        $SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
				        $SPItem["State"] = $Script:RunningState
				        $SPItem.Update()		
                        $Script:ctx.executeQuery()
					}
					else {
<##########
###########
						$SPItem = $Script:FeedAckTable.Items.Add()
						$SPItem["FSN"] = $Script:CurrentACK
						$SPItem["Filename"] = $Result[0].filename
						$SPItem["FileType"] = "R"
						$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
						$SPItem["FSNReusedNum"] = 1

						$SPItem["Title"] = $Result[0].filename
						$Content = Get-Content -encoding Byte -path (Join-Path $Script:NotAck $Result[0].filename)
						$SPItem.attachments.add($Result[0].filename, $Content)

						$SPItem.Update()
###########
##########>
                        $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                        $SPItem = $Script:FeedAckTable.AddItem($listItemInfo)
						$SPItem["FSN"] = $Script:CurrentACK
						$SPItem["Filename"] = $Result[0].filename
						$SPItem["FileType"] = "R"
						$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
						$SPItem["FSNReusedNum"] = 1
						$SPItem["Title"] = $Result[0].filename
                        $SPItem.Update()
                        $Script:ctx.load($Script:FeedAckTable)
                        $Script:ctx.executeQuery()
	                    $listItemInfo = $null

                        # Get the record ID
                        $Script:ctx.load($SPItem)
                        $Script:ctx.executeQuery()

                        $ID = @($SPItem.FieldValues)[0].ID
                        $listName = "FeedAckTable"
                        $fStream=[System.IO.File]::OpenRead(($(Join-Path $Script:NotAck $Result[0].filename)))
                        [String] $fName=$fStream.Name
                        [System.Byte[]]$bytes=New-Object -TypeName System.Byte[] -ArgumentList $fStream.Length
                        [void] $fStream.Read($bytes, 0, [int]$fStream.Length)
                        $fStream.Close()
                        $listWebServiceReference = New-WebServiceProxy -Uri $WSuri -UseDefaultCredential
                        [void] $listWebServiceReference.AddAttachment($listName,$ID,$fName,$bytes)

						$Script:RunningState = 4
<##########
###########
						$SPItem = $StateTable.items[0]
						$SPItem["State"] = $Script:RunningState
						$SPItem.Update()
###########
##########>
	                    $SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
						$SPItem["State"] = $Script:RunningState
	                    $SPItem.Update()
                        $Script:ctx.executeQuery()
					}			
				}	
			}
		}
	}
}

function ReadAck {
<##########
###########
	$QueryCAML = 
@"
		<Where>
				<And>
  					<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
					<Eq><FieldRef Name='FileType' /><Value Type='Text'>R</Value></Eq>
				</And>
		</Where>
"@
	$Query = New-Object Microsoft.Sharepoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:FeedAckTable.GetItems($Query)
	$Query = $null
	$Filename = $Result[0]["Filename"]
##########
#########>

	$QueryCAML = 
@"
        <View>
        <Query>
		<Where>
				<And>
  					<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
					<Eq><FieldRef Name='FileType' /><Value Type='Text'>R</Value></Eq>
				</And>
		</Where>
        </Query>
        </View>
"@
	$Result = GetCAMLResult $Script:FeedAckTable $QueryCAML
	$Filename = @($Result.FieldValues)[0]["Filename"]

	$AllLine = Get-Content (Join-Path $Script:NotAck $Filename)

	# This is ugly as the ack from TRU has a strange line of char after TRAILER line
	$Adjust = 1
	if ($AllLine[-1] -match "^TRAILER") {
		$Adjust = 1
	}
	elseif ($AllLine[-2] -match "^TRAILER") {
		$Adjust = 2
	}
	
	$AllLine[1 .. $($AllLine.length - (1 + $Adjust))] | foreach {
	
#        $AllLine[1 .. $($AllLine.length - 2)] | foreach {
		$Field1 = [double]$_.substring(0,15)
		$Field2 = $_.substring(15,15)
		$Field3 = $_.substring(30,8)
		$Field4 = $_.substring(38,120)
		$Field5 = $_.substring(158,3)
# 7-24-2013
# TRU changes the record spec without following the 161 len
#		$Field6 = $_.substring(161,100)

		$Field6 = $_.substring(161, $_.length -161)
		
		if ("100","101","102","103" -contains $Field5) {
			switch ($Field5) {
				"100" {$Script:RunningState = -5}
				"101" {$Script:RunningState = -6}
				"102" {$Script:RunningState = -7}
				"103" {$Script:RunningState = -8}
			}
		}
		else {
			$Script:RunningState = 5
		}

<##########
###########
		$QueryCAML =
@"
			<Where>
				<And>
  					<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
					<Eq><FieldRef Name='SKN' /><Value Type='Number'>$Field1</Value></Eq>
				</And>
			</Where>
"@
		$Query = New-Object Microsoft.Sharepoint.SPQuery
		$Query.query = $QueryCAML
		$Result = $Script:RecordReceiveTable.GetItems($Query)
		$Query = $null
###########
##########>

		$QueryCAML =
@"
            <view>
            <Query>
			<Where>
				<And>
  					<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
					<Eq><FieldRef Name='SKN' /><Value Type='Number'>$Field1</Value></Eq>
				</And>
			</Where>
            </Query>
            </View>
"@
		$Result = GetCAMLResult $Script:RecordReceiveTable $QueryCAML

		if ($Result.count -ne 0) {
<##########
###########
			$SPItem = $Result[0]
			$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
			$SPItem.Update()
###########
##########>
            $ID = @($Result.FieldValues)[0].ID
        	$SPItem = $Script:RecordReceiveTable.GetItemById($ID)
			$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
			$SPItem.Update()
            $Script:ctx.executeQuery()
		}
		else {
<##########
###########
			$SPItem = $Script:RecordReceiveTable.Items.Add()
###########
##########>

            $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $SPItem = $Script:RecordReceiveTable.AddItem($listItemInfo)

			$SPItem["FSN"] = $Script:CurrentACK
			$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
			$SPItem["SKN"] = $Field1
			$SPItem["Upc"] = $Field2
			$SPItem["Uid"] = $Field3
			$SPItem["Desc"] = $Field4
			$SPItem["ReasonCode"] = $Field5
			$SPItem["ReasonDescr"] = $Field6
#			$SPItem["ReasonDescr"] = $Script:CodeDesc[$Field5]
			$SPItem.Update()

            $Script:ctx.load($Script:RecordSentTable)
            $Script:ctx.executeQuery()
	        $listItemInfo = $null
		}
	}

<##########
###########
	$SPItem = $StateTable.items[0]
	$SPItem["State"] = $Script:RunningState
	$SPItem.Update()
###########
##########>

	$SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
	$SPItem["State"] = $Script:RunningState
	$SPItem.Update()
    $Script:ctx.executeQuery()
}

function UpdateRecordFromAck {
<##########
###########
	$QueryCAML = 
@"
		<Where>
  			<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
		</Where>
"@
	$Query = New-Object Microsoft.Sharepoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:RecordReceiveTable.GetItems($Query)
	$Query = $null
###########
##########>

	$QueryCAML = 
@"
        <View>
        <Query>
		<Where>
  			<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
		</Where>
        </Query>
        </View>
"@
	$Result = GetCAMLResult $Script:RecordReceiveTable $QueryCAML

<##########
###########

	$Result | foreach {
		$QueryCAML = 
@"
			<Where>
  				<Eq><FieldRef Name='SKN_N' /><Value Type='Number'>$($_["SKN"])</Value></Eq>
			</Where>
			<OrderBy>
	   			<FieldRef Name='ID' Ascending='False' />
			</OrderBy>
"@
		$Query = New-Object Microsoft.Sharepoint.SPQuery
		$Query.query = $QueryCAML
		$Result1 = $Script:TruSPList.GetItems($Query)
		$Query = $null

###########
##########>

	$Result.FieldValues | foreach {
		$QueryCAML = 
@"
            <View>
            <Query>
			<Where>
  				<Eq><FieldRef Name='SKN_N' /><Value Type='Number'>$($_["SKN"])</Value></Eq>
			</Where>
			<OrderBy>
	   			<FieldRef Name='ID' Ascending='False' />
			</OrderBy>
            </Query>
            </View>
"@
		$Result1 = GetCAMLResult $Script:TruSPList $QueryCAML
		
		if ($Result1.count -ne 0) {

            $Result1 = @($Result1.FieldValues)

			if ($Result1[0]["FeedBack"] -eq $Script:CodeDesc["1000"]) {

<##########
###########
				$SPItem = $Result1[0]
###########
##########>
                $ID = $Result1[0].ID
	            $SPItem = $Script:TruSPList.GetItemById($ID)
				#$SPItem["FeedBack"] = $_["ReasonDescr"]
				$SPItem["FeedBack"] = $Script:CodeDesc[$_["ReasonCode"]]
				$SPItem["AckFromTRU"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
				$SPItem["UpdateByACK"] = $Script:CurrentACK
				$SPItem.Update()
                $Script:ctx.executeQuery()

<##########
###########
				$SPItem = $_
				$SPItem["ProcessDesc"] = $Script:ReceiveRecordDesc["0"]
				$SPItem.Update()
###########
##########>
                $ID = $_.ID
				$SPItem = $Script:RecordReceiveTable.GetItemById($ID)
				$SPItem["ProcessDesc"] = $Script:ReceiveRecordDesc["0"]
				$SPItem.Update()
                $Script:ctx.executeQuery()
			}
			else {
<##########
###########
			    $SPItem = $_
				$SPItem["ProcessDesc"] = $Script:ReceiveRecordDesc["1"]
				$SPItem.Update()
###########
##########>

                $ID = $_.ID
				$SPItem = $Script:RecordReceiveTable.GetItemById($ID)
				$SPItem["ProcessDesc"] = $Script:ReceiveRecordDesc["1"]
				$SPItem.Update()
                $Script:ctx.executeQuery()
			}
		}
		else {
<##########
###########
			$SPItem = $_
			$SPItem["ProcessDesc"] = $Script:ReceiveRecordDesc["2"]
			$SPItem.Update()
###########
##########>

            $ID = $_.ID
			$SPItem = $Script:RecordReceiveTable.GetItemById($ID)
			$SPItem["ProcessDesc"] = $Script:ReceiveRecordDesc["2"]
			$SPItem.Update()
            $Script:ctx.executeQuery()
		}
	}

<##########
###########

	$QueryCAML =
@"
		<Where>
			<And>
  				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
				<Eq><FieldRef Name='FileType' /><Value Type='Text'>R</Value></Eq>
			</And>
		</Where>

"@
	$Query = New-Object Microsoft.Sharepoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:FeedAckTable.GetItems($Query)
	$Query = $null
###########
##########>

	$QueryCAML =
@"
        <View>
        <Query>
		<Where>
			<And>
  				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
				<Eq><FieldRef Name='FileType' /><Value Type='Text'>R</Value></Eq>
			</And>
		</Where>
        </Query>
        </View>

"@
	$Result = GetCAMLResult $Script:FeedAckTable $QueryCAML

<##########
###########
	$Filename = $Result[0]["Filename"]
###########
##########>

	$Filename = @($Result.FieldValues)[0]["Filename"]

	Move-Item -path (join-path $Script:NotAck $Filename) -destination (Join-Path $Script:Ack $Filename) -force -erroraction silentlycontinue
	
	$Script:RunningState = 3

<##########
###########
	$SPItem = $StateTable.items[0]
	$SPItem["State"] = $Script:RunningState
	$SPItem["AckFSN"] = ++$Script:CurrentACK
	$SPItem.Update()
###########
##########>

	$SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
	$SPItem["State"] = $Script:RunningState
	$SPItem["AckFSN"] = ++$Script:CurrentACK
	$SPItem.Update()
    $Script:ctx.executeQuery()
}

function ExitTruEdi {
	$Script:RunningState = 0

<##########
###########
	$SPItem = $StateTable.items[0]
	$SPItem["State"] = $Script:RunningState
	$SPItem.Update()
###########
##########>

	$SPItem = $Script:StateTable.GetItemById($Script:StateTableID)
	$SPItem["State"] = $Script:RunningState
	$SPItem.Update()
    $Script:ctx.executeQuery()
}

function SendAlert($MsgSubject, $MsgAttachment) {

	$Msg = new-object Net.Mail.MailMessage
	if ($MsgAttachment) {
    	$Att = new-object Net.Mail.Attachment($MsgAttachment)
	}
    $Smtp = new-object Net.Mail.SmtpClient($Script:MailRelay)

    $Msg.From = $Script:MailSender
    $Msg.To.Add($Script:MailRecipient)
    $Msg.Subject = $MsgSubject
    if ($MsgAttachment) {
		$msg.Attachments.Add($att)
	}
	$Smtp.Send($msg)
	if ($MsgAttachment) {
		$Att.dispose()
	}
	
	$Msg = $Att = $Smtp = $null	
}

function SysCriticalErr {
	$QueryCAML =
@"
		<Where>
			<And>
  				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$Script:CurrentACK</Value></Eq>
				<Eq><FieldRef Name='FileType' /><Value Type='Text'>R</Value></Eq>
			</And>
		</Where>

"@
	$Query = New-Object Microsoft.Sharepoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:FeedAckTable.GetItems($Query)
	$Query = $null

	$Filename = $Result[0]["Filename"]
	Move-Item -path (join-path $Script:NotAck $Filename) -destination (Join-Path $Script:Ack $Filename) -force -erroraction silentlycontinue
	SendAlert $Script:StateErrDesc["$Script:RunningState"] $null
	
	$SPItem = $StateTable.items[0]
	$SPItem["AckFSN"] = ++$Script:CurrentACK
	$SPItem.Update()
}

function AckFileErr {
	SendAlert $Script:StateErrDesc["$Script:RunningState"] $(if ($Script:CurrentACKFile) {(Join-Path $Script:NotAck $Script:CurrentACKFile)} else {$null})
	Move-Item -path (Join-Path $Script:NotAck '*') -destination $Script:ErrAck -force -erroraction SilentlyContinue
}

#######################
# Manual resend of feed
#######################

function ResendFeed ($ResendFSN) {
	$QueryCAML =
@"
		<Where>
			<And>
  				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$ResendFSN</Value></Eq>
				<Eq><FieldRef Name='FileType' /><Value Type='Text'>S</Value></Eq>
			</And>
		</Where>
		<OrderBy>
   			<FieldRef Name='FSNReusedNum' Ascending='True' />
		</OrderBy>
"@
	$Query = New-Object Microsoft.Sharepoint.SPQuery
	$Query.query = $QueryCAML
	$Result = $Script:FeedAckTable.GetItems($Query)
	$Query = $null

	if ($Result.count) {
		# create a function to create file name
		
		$ReusedNum = $Result[$Result.count - 1]["FSNReusedNum"]
		
		# Fill in the FeedAck Table
		$FeedFileName = GenFilename
		$SPItem = $Script:FeedAckTable.Items.Add()
		$SPItem["FSN"] = $ResendFSN
		$SPItem["FileType"] = "S"
		$SPItem["FSNReusedNum"] = $ReusedNum + 1
		$SPItem["Filename"] = $FeedFileName
		$SPItem["Date"] = [string]::format("{0:yyyy-MM-ddTHH:mm:ssZ}", [datetime]::utcnow)
		$SPItem.Update()
		
		# Fill in the RecordSent Table
		$QueryCAML = 
@"
		<Where>
			<And>
				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$ResendFSN</Value></Eq>
				<Eq><FieldRef Name='FSNReusedNum' /><Value Type='Number'>$ReusedNum</Value></Eq>
			</And>
    	</Where>
		<OrderBy>
   			<FieldRef Name='SKN' Ascending='True' />
		</OrderBy>

"@
		$Query = New-Object Microsoft.SharePoint.SPQuery
		$Query.query = $QueryCAML
		$Result = $Script:RecordSentTable.GetItems($Query)
		$Query = $null
		
		$Result | foreach {
			$SPItem = $Script:RecordSentTable.Items.Add()
			$SPItem["SKN"] = $_["SKN"]
			$SPItem["FSN"] = $_["FSN"]
			$SPItem["FSNReusedNum"] = $ReusedNum + 1
			$SPItem.Update()
 		}
		
		# Search those records again then pass to outfeed to generate a feed file
		$QueryCAML = 
@"
		<Where>
			<And>
				<Eq><FieldRef Name='FSN' /><Value Type='Number'>$ResendFSN</Value></Eq>
				<Eq><FieldRef Name='FSNReusedNum' /><Value Type='Number'>$($ReusedNum + 1)</Value></Eq>
			</And>
    	</Where>
		<OrderBy>
   			<FieldRef Name='SKN' Ascending='True' />
		</OrderBy>

"@
		$Query = New-Object Microsoft.SharePoint.SPQuery
		$Query.query = $QueryCAML
		$Result = $Script:RecordSentTable.GetItems($Query)
		$Query = $null
	
		# Generate the feed file
		OutFeed $Result $ResendFSN (Join-Path $Script:TempDir $FeedFileName)
		
		# Push to SFTP
		"####"
		"PUSHING TO SFTP..."
		"###"
	
	}
	else {
		"I couldn't find FSN: $ResendFSN in the system, resend unsuccessful!"
	}
}

# Initialize the system
InitializeSys
if (-not $Args.length) {
	Remove-Item -path $Logfile -erroraction SilentlyContinue
	"Enter: " + [datetime]::Now.touniversaltime() | out-file $Logfile -encoding ascii -append
	if ($RunningState -lt 0) {
		$StateErrDesc["$RunningState"]
	}
	else {	
		if ((Check-DropOff) -eq 0) {
			:outloop while ($true) {
				switch ($RunningState) {
					0 {
						FindRecordToSend
		 			}
					1 {
						GenerateFeed
	    				}
					2 {
						PushToSftp
					}
					3 {
						PullFromSftp
					}
					4 {
						ReadAck
					}
					5 {
						UpdateRecordFromAck
					}
					6 {
						ExitTruEdi
						break outloop
					}
					{-5, -6, -7, -8 -contains $_} 
					{
						$StateErrDesc["$RunningState"]
						SysCriticalErr
						break outloop
					}
					{-1, -2, -3, -4 -contains $_}
					{
						$StateErrDesc["$RunningState"]
						AckFileErr
						break outloop
					}
					default { 
						$RunningState
						break outloop
					}
				}
			}
		}
	}
	"Exit: " + [datetime]::Now.touniversaltime() | Out-File $Logfile -encoding ascii -append
	
	if ($SendLog) {
		SendAlert "TRU EDI run time log" $Logfile
	}
}
else {
	ResendFeed $([int]$Args[0])
}
