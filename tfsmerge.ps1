#Created by Bryan Green (codergr33n) 
#Contact: green.c.bryan@gmail.com

#.Net TFS cmdlets setup
# Loads Windows PowerShell snap-in if not already loaded
if ( (Get-PSSnapin -Name Microsoft.TeamFoundation.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PSSnapin Microsoft.TeamFoundation.PowerShell
}

$binpath   = "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0"
    Add-Type -path "$binpath\Microsoft.TeamFoundation.Client.dll"
    Add-Type -Path "$binpath\Microsoft.TeamFoundation.WorkItemTracking.Client.dll"
    Add-Type -Path "$binpath\Microsoft.TeamFoundation.Common.dll"

$server = [Microsoft.TeamFoundation.Client.TeamFoundationServerFactory]::GetServer("https://<tfs url>")
$workItemStore = $server.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
$vcServer = $server.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])


#create the sql command function to get the changesets
$dataSource = "<server name>"
$database = "<database name>"
$sqlCommand = "SELECT DISTINCT C.ChangesetID, 
        WI_Story.System_Id AS UserStoryId
FROM DimChangeset C
JOIN FactCodeChurn FCC
	ON FCC.ChangesetSK = C.ChangesetSK
JOIN DimFile F
	ON FCC.FilenameSK = F.FileSK
JOIN vFactWorkItemChangesetOverlay WIC
	ON C.ChangesetSK = WIC.ChangeSetSK
JOIN DimWorkItem WI_Story
	ON WIC.WorkItemSK = WI_Story.WorkItemSK
	AND WI_Story.System_WorkItemType = 'User Story'
JOIN DimIteration I_Story
	ON WI_Story.IterationSK = I_Story.IterationSK
LEFT JOIN DimPerson P_Story
	ON WI_Story.System_AssignedTo__PersonSK = P_Story.PersonSK
WHERE 
	WI_Story.System_State = 'Test Build' --workitem state
	AND F.FilePath LIKE '$<tfs source path>' + '%'
	--AND I_Story.IterationPath LIKE @iterationPath + '%'
ORDER BY C.ChangesetID ASC"
      

$connectionString = "Data Source=$dataSource; " +
        " User Id=<userid>;Password=<password>; " +
        "Initial Catalog=$database"

$connection = new-object system.data.SqlClient.SQLConnection($connectionString)
$command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
$connection.Open()

$adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
$data = New-Object System.Data.DataSet

#fill the changeset dataset
$adapter.Fill($data)

$connection.Close()

#select the changeset datatable
$table = $data.Tables[0]


#setup the tfs cmd prompt link in order to do the merge later
$tf = get-item "c:\program files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\TF.EXE"
& $tf workspaces /Collection:"https://<TFS URL>" /owner:<useraccount>

#check for changesets
if(!$table)
{
    echo "No changesets to merge."
    exit 0
}

#get latest
& $tf workfold /map "$/<TFS map path>" "D:\<local map path>"
& $tf get "$<tfs map path>" /recursive

#loop through and create the tf merge commands
Foreach ($row in $table.Rows)
{

    #setup the comment and related items
    $changeset = [string]$row.ChangesetID
    $workItem = [string]$row.UserStoryId
    $version  = $changeset + "~" + $changeset
    $comment = $workItem + ": Merge changeset " + $changeset + " to <branch> from <branch>"

    #Write-Host("Working with "  + $changeset)
    # http://youtrack.jetbrains.com/issue/TW-9050
    & $tf changeset $($row.ChangesetID) /noprompt
    # Its very important that the right branch (source of automerge) be listed below. 
    #& $tf merge /candidate "$/Software/Source Code/Test" "$/Software/Source Code/Development" /recursive 
    
    & $tf merge /version:$version  "$/<merge from TFS path>" "$/<merge to TFS path>" /recursive /noprompt 
    
    Write-Host("->Merge Exitcode:" + $LASTEXITCODE);

    if($LASTEXITCODE -eq 0)
    {
        & $tf checkin /noprompt /comment:$comment /override:$comment
    }
    
    Write-Host("->Checkin Exitcode:" + $LASTEXITCODE);
    
    if($LASTEXITCODE -eq 0) #check if the checkin happened
    {

        #get the last checkin to associate it
        $items = Get-TfsItemHistory -HistoryItem "D:\<path to local workspace mapping>"  -Recurse -Stopafter "1" -IncludeItems | Select-Object -Expand "Changes" | Select-Object -Expand "Item"
        $checkinChangesetId = $items[0].ChangesetId;

        #get changeset on
        $changesetAPI = $vcServer.GetChangeset($checkinChangesetId);

        #get workitem
        $workitemAPI = $workItemStore.GetWorkItem($workitem);

        #get linktype
        $linkType = $workItemStore.RegisteredLinkTypes[[Microsoft.TeamFoundation.ArtifactLinkIds]::Changeset];

        Write-Host("->Linking changeset " + $changeset + " to workitem " + $workitem.Title);

        #link the changeset to workitem
        $externalLink = New-Object  Microsoft.TeamFoundation.WorkItemTracking.Client.ExternalLink -ArgumentList $linkType, $changesetAPI.ArtifactUri.AbsoluteUri
        $workItemAPI.Open();
        $workitemAPI.Links.Add($externalLink);
        $workitemAPI.Save();
    }
    
}


#cleanup everything
& $tf undo . /r
exit 0