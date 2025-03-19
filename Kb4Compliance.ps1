<# Documentation Reference:
# Author : Dale Lee
# Revision : Mar 18, 2025 #>
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#Install-module JoinModule -AllowClobber
import-module JoinModule #https://www.powershellgallery.com/packages/JoinModule/
import-Module ImportExcel #https://www.powershellgallery.com/packages/ImportExcel/
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    } #end function
# $Guid=""; #Guid for tracking; generate with new-guid
$ProjInfo = "KnowBe4 API Report Update "
$MyReport = New-Object System.Collections.ArrayList
$set = New-Object System.Collections.ArrayList
$div_rpt_final = New-Object System.Collections.ArrayList
$div_report = New-Object System.Collections.ArrayList
$ActivityLog = New-Object System.Collections.Generic.List[System.Object]

$mail_body = New-Object System.Collections.ArrayList
$mail_fr = 'from@example.com'; $mail_to = 'to@example.com'; $PSEmailServer = 'mail.example.com' ; 

$reportloc = 'C:\proj\kb4rpt\'
$reportname = 'KB4MasterReport.xlsx'
$kb4apitoken = 'knowbe4apitoken'

#region - function get campaigns
    function get-campaigns{
        Param([Parameter(Mandatory=$true)][string]$URL,
            [Parameter(Mandatory=$false)][object] $kb4apitoken)
        $kb4_auth = @{
            "ContentType" = 'application/json'
            "Method"      = 'GET'
            "Headers"     = @{
                "Authorization" = "Bearer $($kb4apitoken)"
                "Accept"        = "application/json"
                }
            }
        $campaigns = Invoke-RestMethod -Uri $URL @kb4_auth
        Return $campaigns
        }   
    #endregion - function get campaigns
#region - function get active users
    function get-ActiveUsers{
        Param([Parameter(Mandatory=$false)][string]$URL,
            [Parameter(Mandatory=$true)][object] $kb4apitoken,
            [Parameter(Mandatory=$false)][int]$pagesize = 500
            )
        $kb4_auth = @{
            "ContentType" = 'application/json'
            "Method"      = 'GET'
            "Headers"     = @{
                "Authorization" = "Bearer $($kb4apitoken)"
                "Accept"        = "application/json"
                }
            }
       #$users = Invoke-RestMethod -Uri $URL @kb4_auth
        $ActiveUsers = New-Object System.Collections.Generic.List[System.Object];$page=0;
            do {$page++ ; write-host $page ; write-host $pagesize
                $URL = 'https://us.api.knowbe4.com/v1/users/'+"?page=$($page)"+"&per_page=$($pagesize)"
                $ReturnedUser = Invoke-RestMethod -Uri $URL @kb4_auth -Verbose
                $ActiveUsers += $ReturnedUser ; write-host $ActiveUsers.count $ReturnedUser.count
                }while ($ReturnedUser.count -gt 0)
        Return $ActiveUsers
        }  
    #endregion - function get active users
#region - function get enrollments
    function get-Enrollments{
        Param([Parameter(Mandatory=$false)][string]$URL,
            [Parameter(Mandatory=$true)][object] $kb4apitoken,
            [Parameter(Mandatory=$false)][int]$pagesize = 500
            )
        $kb4_auth = @{
            "ContentType" = 'application/json'
            "Method"      = 'GET'
            "Headers"     = @{
                "Authorization" = "Bearer $($kb4apitoken)"
                "Accept"        = "application/json"
                }
            }
       #$users = Invoke-RestMethod -Uri $URL @kb4_auth
        $Enrollments = New-Object System.Collections.Generic.List[System.Object];$page=0;
            do {$page++ ; write-host $page ; write-host $pagesize
                $URL = 'https://us.api.knowbe4.com/v1/training/enrollments/'+"?page=$($page)"+"&per_page=$($pagesize)";
                $ReturnedEnrollments = Invoke-RestMethod -Uri $URL @kb4_auth -Verbose
                $Enrollments+=$ReturnedEnrollments; write-host $Enrollments.count $ReturnedEnrollments.count
            }while ($ReturnedEnrollments.count -gt 0)
        Return $Enrollments
        }  
    #endregion - function get enrollments
 
#region if the knowbe4 master report is < 20 hours old import from that file, else update from the knowbe4 api
    $maxage=20; $bar=(get-date).AddHours(-$maxage)
    $ActivityLog.add("$(Get-TimeStamp) Checking date of kb4 report file.")
    
    try{$filetime= get-itemproperty -path "$($reportloc)$($reportname)" -Name LastWriteTime
        $ActivityLog.Add("$(Get-TimeStamp) file has date[$($filetime)].")
    }catch{
        $filetime = $bar-1
        $ActivityLog.Add("$(Get-TimeStamp) file not found; setting filetime to null")
    }
     if($filetime.LastWriteTime -ge $bar){
        $ActivityLog.Add("Skipping KnowBe4 API query and importing data from previous run.")
        $Exported_Campaigns=Import-Excel -path "$($reportloc)$($reportname)" -WorksheetName Campaigns
        $Exported_Users=Import-Excel -path "$($reportloc)$($reportname)" -WorksheetName Users
        $KB4_Users = $Exported_Users
        $EENR=Import-Excel -path "$($reportloc)$($reportname)" -WorksheetName Enrollments
        $KB4_Enrollments=$EENR
        }
    elseif($filetime.LastWriteTime -lt $bar){
        $ActivityLog.Add("Performing KnowBe4 API Query and updating File")
       #region Get all campaigns
            $URL='https://us.api.knowbe4.com/v1/training/campaigns/'
            $KB4_Campaigns=get-campaigns -URL $URL -kb4apitoken $kb4apitoken
            $Exported_Campaigns =New-Object System.Collections.Generic.List[System.Object]
            foreach ($CreatedCampaign in $KB4_Campaigns){
                $c_groups=@(); foreach($g1 in $CreatedCampaign.groups){   $c_groups+= "$($g1.name) [$($g1.group_id)]"}
                $c_content=@();foreach($c1 in $CreatedCampaign.content){$c_content+="$($c1.name) [$($c1.store_purchase_id)]"}
                if([string]::IsNullOrWhitespace($CreatedCampaign.start_date)){$T1=""}else{$T1=(([DateTime]$CreatedCampaign.start_date).ToUniversalTime())}
                if([string]::IsNullOrWhitespace($CreatedCampaign.end_date)){$T2=""}else{$T2=(([DateTime]$CreatedCampaign.end_date).ToUniversalTime())}
                $Exported_Campaigns.add([pscustomobject]@{
                    campaign_id=[int]$CreatedCampaign.campaign_id
                    name=$CreatedCampaign.name
                    groups=($c_groups)-join", "
                    status=$CreatedCampaign.status
                    content=($c_content)-join", "
                    duration_type=$CreatedCampaign.duration_type
                    start_date= $T1
                    end_date= $T2
                    relative_duration=$CreatedCampaign.relative_duration
                    completion_percentage=$CreatedCampaign.completion_percentage
                    })
                }
            $Exported_Campaigns|Export-Excel -path "$($reportloc)$($reportname)" -WorksheetName Campaigns -TableName Campaigns -TableStyle Light2 -BoldTopRow -FreezeTopRow -ClearSheet -autosize
            $ActivityLog.Add("exported campaigns to excel doc.")
            #endregion

        #region get active users
            
            $KB4_Users = get-activeUsers -kb4apitoken $kb4apitoken -pagesize 499
            $Exported_Users=New-Object System.Collections.Generic.List[System.Object]
                foreach ($user in $KB4_Users){
                    if([string]::IsNullOrWhitespace($user.joined_on          )){$uj=""}else{$uj=(([DateTime]$user.joined_on          ).ToUniversalTime())}
                    if([string]::IsNullOrWhitespace($user.last_sign_in       )){$ul=""}else{$ul=(([DateTime]$user.last_sign_in       ).ToUniversalTime())}
                    if([string]::IsNullOrWhitespace($user.employee_start_date)){$ue=""}else{$ue=(([DateTime]$user.employee_start_date).ToUniversalTime())}
                    if([string]::IsNullOrWhitespace($user.archived_at        )){$ua=""}else{$ua=(([DateTime]$user.archived_at        ).ToUniversalTime())}
                    $Exported_Users.add([pscustomobject]@{
                        id=$user.id
                        employee_number=$user.employee_number
                        first_name=$user.first_name
                        last_name=$user.last_name
                        job_title=$user.job_title
                        email=$user.email
                        phish_prone_percentage=$user.phish_prone_percentage
                        phone_number=$user.phone_number
                        extension=$user.extension
                        mobile_phone_number=$user.mobile_phone_number
                        location=$user.location
                        division=$user.division
                        manager_name=$user.manager_name
                        manager_email=$user.manager_email
                        provisioning_managed=$user.provisioning_managed
                        provisioning_guid=$user.provisioning_guid
                        groups=($user.groups)-join "|"
                        current_risk_score=$user.current_risk_score
                        aliases=($user.aliases)-join ";"
                        joined_on=$uj
                        last_sign_in= $ul
                        status	=$user.status
                        organization	=$user.organization
                        department	=$user.department
                        language	=$user.language
                        comment	=$user.comment
                        employee_start_date	= $ue
                        archived_at	=$ua
                        custom_field_1	=$user.custom_field_1
                        custom_field_2	=$user.custom_field_2
                        custom_field_3	=$user.custom_field_3
                        custom_field_4	=$user.custom_field_4
                        custom_date_1	=$user.custom_date_1
                        custom_date_2=$user.custom_date_2
                        })
                    }
            $Exported_Users|Export-Excel -path "$($reportloc)$($reportname)" -WorksheetName Users -TableName Users -TableStyle Light2 -BoldTopRow -FreezeTopRow -ClearSheet -autosize
            $ActivityLog.Add("$($Exported_Users.count) exported users to excel doc.")
            #endregion get active users 

        #region Get enrollments
            $KB4_Enrollments=New-Object System.Collections.Generic.List[System.Object];

            $EENR = get-enrollments -kb4apitoken $kb4apitoken
            foreach ($e in $EENR){
                if([string]::IsNullOrWhitespace($e.enrollment_date)){$ee=""}else{$ee=(([DateTime]$e.enrollment_date).ToUniversalTime())}
                if([string]::IsNullOrWhitespace($e.start_date     )){$es=""}else{$es=(([DateTime]$e.start_date     ).ToUniversalTime())}
                if([string]::IsNullOrWhitespace($e.completion_date)){$ec=""}else{$ec=(([DateTime]$e.completion_date).ToUniversalTime())}
                $KB4_Enrollments.add([pscustomobject]@{
                        enrollment_id=$e.enrollment_id
                        content_type=$e.content_type
                        module_name=$e.module_name
                        id=$e.user.id
                        first_name=$e.user.first_name
                        last_name=$e.user.last_name
                        email=$e.user.email
                        campaign_name =$e.campaign_name
                        enrollment_date = $ee
                        start_date = $es
                        completion_date =$ec
                        status =$e.status
                        time_spent =$e.time_spent
                        policy_acknowledged =$e.policy_acknowledged
                        score =$e.score
                        })
                }
            $KB4_Enrollments|Export-Excel -path "$($reportloc)$($reportname)" -WorksheetName Enrollments -TableName Enrollments -TableStyle Light2 -BoldTopRow -FreezeTopRow -ClearSheet -autosize
            $ActivityLog.Add("exported [$($KB4_Enrollments.count)] enrollments to excel doc.")
            #endregion get enrollments
     
        }
    #endregion

#region Calculate enrollments by user
    $enrollments_all = $KB4_Enrollments|Where-Object{$_.status -ne "" -and $_.campaign_name -ne ""}
    $ActivityLog.Add("$(Get-TimeStamp) filtering enrollments...complete.<BR>")
    $enrollments_GrpByUser = $enrollments_all|group-object -property id
    $ActivityLog.Add("$(Get-TimeStamp) grouping enrollments...complete.<BR>")
    $enrollments_bucket=@(); #-Property @{id=0}
    
    $compl_sort_custom = "Passed","Past Due","In Progress","Not Started"
    foreach($grouped in $enrollments_GrpByUser){
        #each user may have multiple enrollments; one user per multiple campaings per multiple trainings for each campaign 
        #gathered enrollments will represent the multiple enrollments for each user, which gets reset for each users' grouped enrollments
        #$gathered_enrollments=New-Object System.Collections.ArrayList;# -Property @{id=0}
        $gathered_enrollments=@(); #-Property @{id=0}
        $user_enrollments_by_campaign = $grouped.group|Group-Object -Property campaign_name # < grouped by campaign
        #write-host "-$($user_enrollments_by_campaign.Group[0].email) $($user_enrollments_by_campaign.Group[0].campaign_name)"
        foreach ($foo in $user_enrollments_by_campaign){ # < loops through each campgain
            $single_enrollment = @();#New-Object System.Collections.ArrayList;
            $sorted_enrollment=($foo.group |Sort-Object {$compl_sort_custom.IndexOf($_.Result)})[0]

            if(($NULL -eq $sorted_enrollment.id) -or ($NULL -eq $sorted_enrollment.campaign_name)){
                #skip
            }elseif($sorted_enrollment.status -eq "Passed"){
                $single_enrollment = New-Object -TypeName PSCustomObject -Property @{
                id=$sorted_enrollment.id;$sorted_enrollment.campaign_name = [DateTime]$sorted_enrollment.completion_date}
            }elseif($sorted_enrollment.status -eq "Past Due"){
                $single_enrollment = New-Object -TypeName PSCustomObject -Property @{
                id=$sorted_enrollment.id;$sorted_enrollment.campaign_name = "Past Due"}
            }elseif($sorted_enrollment.status -eq "In Progress"){
                $single_enrollment = New-Object -TypeName PSCustomObject -Property @{
                id=$sorted_enrollment.id;$sorted_enrollment.campaign_name = "In Progress"}
            }elseif($sorted_enrollment.status -eq "Not Started"){
                $single_enrollment = New-Object -TypeName PSCustomObject -Property @{
                id=$sorted_enrollment.id;$sorted_enrollment.campaign_name = "Not Started"}
            } else{
            } #end if
            
            if($single_enrollment.count -eq 0){#skip
            }elseif($gathered_enrollments.count -eq 0){$gathered_enrollments=$single_enrollment
                #write-host  " 294 | A[$($single_enrollment.count)] [$($gathered_enrollments.count)] [$($enrollments_bucket.count)]"
            }else{$gathered_enrollments=Merge-Object -LeftObject $gathered_enrollments -Property id -RightObject $single_enrollment -JoinType Full  
 
                #write-host  "298 | C[$($single_enrollment.count)] [$($gathered_enrollments.count)] [$($enrollments_bucket.count)]"
            } #endif single_enrollment
           # write-host  "ID[$($sorted_enrollment.id)] A[$($single_enrollment.count)] [$($gathered_enrollments.count)] [$($enrollments_bucket.count)]"
          
        }
        #end foreach
        if($gathered_enrollments.count -eq 0){
            #skip
        }elseif($enrollments_bucket.count -eq 0){
            $enrollments_bucket=$gathered_enrollments
           # write-host  "308 | A[$($single_enrollment.count)] [$($gathered_enrollments.count)] [$($enrollments_bucket.count)]"
        #skip
        }else{
            #write-host  "ID[$($sorted_enrollment.id)] 311 C[$($single_enrollment.count)] [$($gathered_enrollments.count)] [$($enrollments_bucket.count)]"
            $enrollments_bucket=$enrollments_bucket| Merge-Object $gathered_enrollments -On id
           # $enrollments_bucket=Merge-Object -LeftObject $enrollments_bucket -Property id -RightObject $gathered_enrollments  -JoinType Full
          
           # $enrollments_bucket = $enrollments_bucket|merge($gathered_enrollments); 
            #write-host "315 $($enrollments_bucket.count) $($gathered_enrollments.count)"
        } #endif single_enrollment
    }
    #endregion
    $ActivityLog.Add("$(Get-TimeStamp) Merging users and enrollments.<BR>")
    $ActivityLog.Add("$(Get-TimeStamp) KB4_Users[0]: $($KB4_Users[0]).<BR>")
    $ActivityLog.Add("$(Get-TimeStamp) Enrollments_bucket[0]: $($enrollments_bucket[0]).<BR>")
    #$UsersbyCompletion=join-object -Left $KB4_Users -right $enrollments_bucket -LeftJoinProperty id -RightJoinProperty id
    $UsersbyCompletion=$KB4_Users| merge-object $enrollments_bucket -On id
    $ActivityLog.Add("$(Get-TimeStamp) Completed merge.<BR>")
  #region export usersbycompletion
    $ActivityLog.Add("$(Get-TimeStamp) Exporting Users By Completion...<BR>")
    $UsersbyCompletion|Export-Excel -path "$($reportloc)$($reportname)" -WorksheetName UsersByCompletion -TableStyle Light2 -TableName UsersByCompletion -BoldTopRow -FreezeTopRow -ClearSheet -autosize
    $ActivityLog.Add("$(Get-TimeStamp) Completed Export.<BR>")
    #endregion

#region send report
    $mail_body.Add("<B>KnowBe4 API Report Update.<B>")
    $mail_body.Add("<Blockquote>$($div_rpt_final|ConvertTo-PSHTMLTable)</Blockquote><P>")
    $mail_body.Add("KnowBe4 Training Report [$($sourcefile_info.Name)] was last obtained on [$($sourcefile_info.LastWriteTime)].<P>")
    $mail_body.Add($ActivityLog)
    #$FinishTime = (Get-Date).ToLocalTime(); $mail_body.Add("Report Finished at $FinishTime `r"); $mail_body.Add("[nosig]")
    $reportpath = "C:\proj\kb4rpt\"
    $B_List = (Get-ChildItem -Path $ReportPath -Filter "KB4MasterReport*").fullname
    $ProjInfo+=" C:$($Exported_Campaigns.count) U:$($Exported_Users.count) E:$($KB4_Enrollments.count)"

    Send-MailMessage -From $mail_fr -To $mail_to -Cc $mail_fr -Subject $ProjInfo -BodyAsHtml ($mail_body|Out-String) -Attachments $B_List 
    #endregion
