function Start-Main {
$clientid = 'CLIENTID'
$tenantid = 'TENANTID'

$token = GetDelegatedGraphToken -clientid $clientid -tenantid $tenantid
$groups = listgroups -token $token
foreach ($group in $groups) {
    $plans = listplans -groupid $group.id -token $token
    if ($plans.count -ne 0) {
        foreach ($plan in $plans) {
            exportplanner -token $token -planid $plan.id -groupid $plan.container.containerid
        }
    }
}
}
function GetDelegatedGraphToken {

    <#
    .SYNOPSIS
    Azure AD OAuth Application Token for Graph API
    Get OAuth token for a AAD Application using delegated permissions via the MSAL.PS library(returned as $token)

    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy

    .PARAMETER redirectURI
    -is the redirectURI specified in the application registration, default value is https://localhost

    #>

    # Application (client) ID, tenant ID and secret
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $clientID,
        [parameter(Mandatory = $true)]
        [String]
        $tenantID,
        [parameter(Mandatory = $false)]
        $RedirectURI = "https://localhost"
    )

    $Token = Get-MsalToken -DeviceCode -ClientId $clientID -TenantId $tenantID -RedirectUri $RedirectURI

    return $token
}


#DON'T USE THIS ANYMORE
function RunQueryandEnumerateResults {
    <#
    .SYNOPSIS
    Runs Graph Query and if there are any additional pages, parses them and appends to a single variable
    
    .PARAMETER apiUri
    -APIURi is the apiUri to be passed
    
    .PARAMETER token
    -token is the auth token
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $apiUri,
        [parameter(Mandatory = $true)]
        $token

    )

    #Run Graph Query
    write-host running $apiuri -foregroundcolor blue
    $results = ""
    $statuscode = ""

    do {
        $results = ""
        $statuscode = ""
    do {
        try{
            $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
            $statusCode = $results.statuscode
        } catch {
            $StatusCode = $_.exception.Response.StatusCode.value__

            if ($statuscode -eq 429) {
                Write-Warning "Got throttled by MS. Sleeping for 45s"
                Start-Sleep -seconds 45
            }
            else {
                write-error $_.Exception
            }
        }
    } while ($statuscode -eq 429)

        if ($results.Value) {
            $resultsValue += $results.value
        } else {
            $resultsValue += $results
        }
        $apiuri = $results.'@odata.nextlink'
    } until (!($apiuri))

    #Output Results for debug checking
    #write-host $results

    ##Return completed results
    return $ResultsValue 
}

function ListGroups {
    <#
    .SYNOPSIS
    Runs Graph Query to list groups in the tenant
    
    .PARAMETER token
    -token is the auth token
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $false)]
        $SearchTerm
    )
    ##Gets Unified Groups

    if ($SearchTerm) {
        $apiUri = "https://graph.microsoft.com/v1.0/groups/?`$filter=groupTypes/any(c:c+eq+'Unified') and startsWith(mail,'$SearchTerm')"
        write-host $apiuri
        $Grouplist = RunQueryandEnumerateResults -token $token.accesstoken -apiUri $apiUri   
    }
    else {
    
        $apiUri = "https://graph.microsoft.com/beta/groups/?`$filter=groupTypes/any(c:c+eq+'Unified')"
        $Grouplist = RunQueryandEnumerateResults -token $token.accesstoken -apiUri $apiUri
    }
    Write-host Found $grouplist.count Groups to process -foregroundcolor yellow

    Return $Grouplist
}

function ListPlans {
    <#
    .SYNOPSIS
    Runs Graph Query to list groups in the tenant
    
    .PARAMETER token
    -token is the auth token

    .PARAMETER GroupID
    -the GroupID of the group continaing the plan
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $false)]
        $GroupID
    )

    $apiUri = "https://graph.microsoft.com/beta/groups/$($Groupid)/planner/plans"
    $Plans = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken

    Return $plans
}

function TestCreateFolder {
    Param(
        [parameter(Mandatory = $true)]
        $directoryPath 
    )
    if(!(Test-Path -path $directoryPath))  
    {  
     New-Item -ItemType directory -Path $directoryPath
     Write-Host "Folder path has been created successfully at: " $directoryPath 
     }
    else 
    { 
    Write-Debug "The given folder path $directoryPath already exists"; 
    }
}

function ListUsers {
    <#
    .SYNOPSIS
    Runs Graph Query to create a hash table mapping users' IDs to displayNames, because for some reason the tasks api is stupid and doesn't always populate the displayName.
    
    .PARAMETER token
    -token is the auth token
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $false)]
        $SearchTerm
    )
    ##Gets Unified Groups

    if ($SearchTerm) {
        $apiUri = "https://graph.microsoft.com/v1.0/users/?`$filter=startsWith(mail,'$SearchTerm')"
        write-host $apiuri
        $Userlist = RunQueryandEnumerateResults -token $token.accesstoken -apiUri $apiUri   
    }
    else {
    
        $apiUri = "https://graph.microsoft.com/beta/users/"
        $Userlist = RunQueryandEnumerateResults -token $token.accesstoken -apiUri $apiUri
    }
    $outputtable = @{}
    foreach ($user in $userlist){
        $outputtable.add($user.id,$user.displayName)
    }
    Write-host Found $grouplist.count Groups to process -foregroundcolor yellow

    Return $outputtable
}

function exportplanner {
    <#
    .SYNOPSIS
    This function gets Graph Token from the GetGraphToken Function and uses it to request a new guest user

    .PARAMETER token
    -is the source auth token
    
    .PARAMETER PlanID
    -is the Plan ID of the source Plan
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $PlanID,
        [parameter(Mandatory = $true)]
        $GroupID 
    )

    $defaultcategories = get-content C:\plannermigrator\DefaultCategories.json | ConvertFrom-Json

    $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($planid)/"
    $Plan = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -Uri $apiUri -Method Get)
    $planname = $plan.title
    $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($planid)/details"
    $PlanDetails = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -Uri $apiUri -Method Get)
    
    $PlanDetailsExport = [PSCustomObject]@{
        categoryDescriptions = $PlanDetails.categoryDescriptions
    }
    $directoryPath = "c:\plannermigrator\exportdirectory\$($planname)\"
    TestCreateFolder -directoryPath $directoryPath
   
    $PlanDetailsExport  | convertto-json -depth 10 |  out-file "c:\plannermigrator\exportdirectory\$($planname)\$($planid)-planDetails.json" -NoClobber -Append

    $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($planid)/buckets"
    $buckets = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
    $bucketsTable = @{}

    $buckets | ForEach-Object {$bucketstable[$_.id]=$_.name}
    if ($buckets) {
        $directoryPath = "c:\plannermigrator\exportdirectory\$($planname)\"
        TestCreateFolder -directoryPath $directoryPath
        $buckets | convertto-json -depth 10 |  out-file "c:\plannermigrator\exportdirectory\$($planname)\$($planid)-buckets.json" -NoClobber -Append
    }   
    
    $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($planid)/tasks"
    $exporttasks = @()
    $tasks = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
    if ($tasks.count -ne 0) {
        $directoryPath = "c:\plannermigrator\exportdirectory\$($planname)\"
        TestCreateFolder -directoryPath $directoryPath
        $tasks  | convertto-json -depth 10 |  out-file "c:\plannermigrator\exportdirectory\$($planname)\$($planid)-tasks.json" -NoClobber -Append
        $i = 0
        $count = $tasks.count
        foreach ($task in $tasks) {
            write-progress -activity "Exporting Tasks" -percentcomplete (($i/$count)*100)
            $apiUri = "https://graph.microsoft.com/beta/planner/tasks/$($task.id)/details"
            $taskdetails = RunQueryandEnumerateResults -token $token.accesstoken -apiuri $apiuri
            #$taskdetails = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -Uri $apiUri -Method Get)
            $directoryPath = "c:\plannermigrator\exportdirectory\$($planname)\"
            TestCreateFolder -directoryPath $directoryPath
            $taskdetails  | convertto-json -depth 10 |  out-file "c:\plannermigrator\exportdirectory\$($planname)\$($task.id)-taskdetails.json" -NoClobber -Append
            if($null -ne $task.conversationthreadid){
                $apiUri = "https://graph.microsoft.com/beta/groups/$($GroupID)/threads/$($task.conversationThreadId)/posts"
                $posts = RunQueryandEnumerateResults -apiuri $apiUri -token $token.accesstoken

                $body = ""

                foreach ($post in $posts) {
                    $sender = $post.from.emailaddress.name
                    $sentTime = get-date -date $post.receivedDateTime
                    $body=$body + "<p>Sent By: $($sender) on $($sentTime)</p><hr>"
                    $html = New-Object -ComObject "HTMLFile"
                    $html.IHTMLDocument2_write($post.body.content)
                    $body =$body + $html.all.tags("div")[0].innerhtml
                    $body = $body + "<hr><br>"

                }
                $title = "Comments on " + $task.title

                $directoryPath = "c:\plannermigrator\exportdirectory\$($planname)\comments\"
                TestCreateFolder -directoryPath $directoryPath
                convertto-html -body $body -title $title | out-file "C:\plannermigrator\exportdirectory\$($planname)\comments\$($task.id)-comments.html"
            }


            #might need to to ifs to iterate over null refs?
            if($null -ne $task.id){$taskid = $task.id} else {$taskid = ""}
            if($null -ne $task.title){$tasktitle = $task.title} else {$tasktitle = ""}
            if($null -ne $taskdetails.description){$taskdescription = $taskdetails.description} else {$taskdescription = ""}
            if($null -ne $task.percentComplete){$taskpercentComplete = $task.percentComplete} else {$taskpercentComplete = ""}
            if($null -ne $task.priority){$taskpriority = $task.priority} else {$taskpriority = ""}
            if($null -ne $task.createddatetime){$taskcreateddate = get-date $task.createdDateTime -format MM/dd/yyyy} else {$taskcreateddate = ""}
            if($null -ne $task.startdatetime){$taskstartdate = get-date $task.startDateTime -format MM/dd/yyyy} else {$taskstartdate = ""}
            if($null -ne $task.duedatetime){$taskduedate = get-date $task.dueDateTime -format MM/dd/yyyy} else {$taskduedate = ""}
            if($null -ne $task.completeddatetime){$taskcompleteddate = get-date $task.completedDateTime -format MM/dd/yyyy} else {$taskcompleteddate = ""}
            if($null -ne $task.bucketid){$bucketname = $bucketstable[$task.bucketid]} else {$bucketname = ""}
            $assignments = ""
            foreach($assignment in ($task.assignments | get-member -type noteproperty | Select-Object -expandproperty name)){
                $assignments = $assignments + $usertable[$assignment] + "`n"
            }
            $attachments = ""
            foreach ($attachment in ($taskdetails.references | get-member -type noteproperty | Select-Object -expandproperty name)){
                $attachments = $attachments + [System.Web.HttpUtility]::UrlDecode($attachment) + "`n"
            }
            $checklists = ""
            if($null -ne $taskdetails.checklist){
                foreach ($check in ($taskdetails.checklist | get-member -type noteproperty)) {
                    $checkobj = $taskdetails.checklist.($check.name)
                    if ($checkobj.isChecked){$checklists = $checklists + "Completed - ";}
                    $checklists = $checklists + $checkobj.title + "`n"
                }
            }
            if($null -ne $task.createdby){$taskcreatedby = $usertable[$task.createdBy.user.id]} else {$taskcreatedby = ""}
            if($null -ne $task.completedby){$taskcompletedby = $usertable[$task.completedBy.user.id]} else {$taskcreatedby = ""}
            #Comments thread
            if($null -ne $task.conversationThreadId){$taskthread = "=HYPERLINK(`"comments\" + $task.id + "-comments.html`")"} else {$taskthread = ""}
            $labels = ""
            foreach($category in ($task.appliedcategories | get-member -type noteproperty | Select-Object -expandproperty name)){
                if($plandetails.categoryDescriptions.$category -ne ""){
                    $labels = $labels + $plandetails.categoryDescriptions.$category + "`n"
                } else {
                    $labels = $labels + $defaultcategories.categoryDescriptions.$category + "`n"
                }
            }


            $exporttask = [PSCustomObject]@{
                taskID = $taskid
                taskName = $tasktitle
                description = $taskdescription
                bucketName = $bucketname
                progress = $taskpercentComplete
                Priority = $taskPriority
                assignedTo = $assignments
                createdBy = $taskcreatedby
                createdDate = $taskcreateddate
                startDate = $taskstartdate
                dueDate = $taskduedate
                completedDate = $taskcompleteddate
                completedBy = $taskcompletedby
                threadid = $taskthread
                attachments = $attachments
                #threadExport
                checklist = $checklists
                labels = $labels
            }

            $exporttasks = $exporttasks + $exporttask
            $i++;
            #start-sleep -m 750
        }
        #write-host $exporttasks
        $exporttasks | Export-excel -path "c:\plannermigrator\exportdirectory\$($planname)\$($planname).xlsx" -autosize -autonamerange -worksheetname Plan
    }
}

Start-Main


