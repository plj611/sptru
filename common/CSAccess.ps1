# README: Need to change siteURL and TruSPList

$Script:siteURL = "https://teams./packaging/tru/"

$Script:WSuri = "https://teams./packaging/tru/_vti_bin/Lists.asmx?wsdl"
$creds = New-Object System.Management.Automation.PsCredential("user", (ConvertTo-SecureString "password" -AsPlainText -Force))
$ClientDLL = [string]::Format("{0}\{1}", $PSScriptRoot, "Microsoft.SharePoint.Client.dll")
$RuntimeDLL = [string]::Format("{0}\{1}", $PSScriptRoot, "Microsoft.SharePoint.Client.Runtime.dll")

Add-Type -Path $ClientDLL
Add-Type -Path $RuntimeDLL
$Ctx =  New-Object Microsoft.SharePoint.Client.ClientContext($Script:siteURL)
$Ctx.credentials = $creds

function GetCAMLResult($list, $caml) {

#    $list=$Script:Ctx.Web.Lists.GetByTitle($listname)
    	
    $cquery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $cquery.ViewXml=$caml

    $listItems = $list.GetItems($cquery)
    $ctx.Load($listItems)
    $ctx.ExecuteQuery()

    $cquery = $null

	return $listItems
}

function InitializeSys() {
    # Initialize the lists
    
    #$Script:TruSPList = $Script:Ctx.Web.Lists.GetByTitle("Verification_dev")
	$Script:TruSPList = $Script:Ctx.Web.Lists.GetByTitle("Item_Verification")
    $Script:StateTable = $Script:Ctx.Web.Lists.GetByTitle("StateTable")
    $Script:FeedAckTable = $Script:Ctx.Web.Lists.GetByTitle("FeedAckTable")
    $Script:RecordSentTable = $Script:Ctx.Web.Lists.GetByTitle("RecordSentTable")
    $Script:RecordReceiveTable = $Script:Ctx.Web.Lists.GetByTitle("RecordReceiveTable")

    # Initialize the system state, FSN and ACK
    $caml="<View><RowLimit>1</RowLimit></View>"
    $Result = GetCAMLResult $Script:StateTable $caml
    $Script:CurrentFSN = $Result.FieldValues["CurrentFSN"]
    $Script:CurrentACK = $Result.FieldValues["AckFSN"]
    $Script:RunningState = $Result.FieldValues["State"]
    $Script:StateTableID = $Result.ID
}
