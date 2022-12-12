# PIM API added as the graph API does not support PAG and alerts. 
# privilegedAccess, privilegedAccess/azureResources/resources

# get PAGs
function get-PIMGroups
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false)]
        [String]$groupid,
        [Parameter(Mandatory=$false)]
        [String]$groupname
    )
    Process
    {

        $API="privilegedAccess/aadGroups/resources"

        If (![string]::IsNullOrEmpty($groupid)){
            $queryString = "`$filter=id eq '$groupid'"
        } elseIf (![string]::IsNullOrEmpty($groupname)){
            $queryString = "`$filter=displayname eq '$groupname'"          
        } else {
            $queryString=$NULL
        }

        $results=Call-MSPIMAPI -AccessToken $AccessToken -API $API -QueryString $queryString
        
        if ($results) { 
            return $results | select id,displayName,type,externalId,status
        } else {
            return $NULL
        }

    }
}


# get PAG role definitions
function get-PIMGrouprolesettings
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(ParameterSetName='groupid',Mandatory=$True)]
        [String]$groupid,
        [Parameter(ParameterSetName='groupname',Mandatory=$True)]
        [String]$groupname,
        [Parameter(Mandatory=$false)]
        [String]$rolename

    )
    Process
    {
               
        $API="privilegedAccess/aadGroups/roleSettingsv2"

        $querystring="`$expand=resource,roleDefinition(`$expand=resource)"
        
        If (![string]::IsNullOrEmpty($groupid)){
            $queryString = "$querystring&`$filter=(resource/id eq '$groupid')"
        } elseIf (![string]::IsNullOrEmpty($groupname)){

            $paggroup = get-PIMGroups -AccessToken $AccessToken -groupname $groupname 
            if ($paggroup) {
                $queryString = "$querystring&`$filter=(resource/id eq '$($paggroup.id)')"          
            } else {

                write-error "No PAG found with name $groupname"
                return $null
            }
           
        } 

        If (![string]::IsNullOrEmpty($rolename)){

            $queryString = "$querystring and (roleDefinition/displayName eq '$rolename')"
        }

        $results=Call-MSPIMAPI -AccessToken $AccessToken -API $API -QueryString $queryString
        
        if ($results){
            $rolesettings = @()

            foreach($result in $results){
               if (![string]::IsNullOrEmpty($result.Id)) { 
  
                
                    $activation = @{}
                    $Eligible = @{}
                    $Active = @{}

                    foreach ($lifeCycleManagement in $result.lifeCycleManagement) {

                        # activation settings
                        if (($lifeCycleManagement.caller -eq 'EndUser') -and ($lifeCycleManagement.level -eq 'Member')) {

                                foreach($value  in  $lifeCycleManagement.value) {
                                    if ($value.ruleIdentifier -eq 'ExpirationRule') {
                                        $Activation['ActivationDuration'] = $($value.setting | convertfrom-json).maximumGrantPeriodInMinutes
                                    }

                                    if ($value.ruleIdentifier -eq 'MfaRule') {
                                        $Activation['mfaRequired'] = $($value.setting | convertfrom-json).mfaRequired
                                    }

                                    if ($value.ruleIdentifier -eq 'JustificationRule') {
                                        $Activation['JustificationRequired'] = $($value.setting | convertfrom-json).required
                                    }

                                    if ($value.ruleIdentifier -eq 'TicketingRule') {
                                        $Activation['ticketingRequired'] = $($value.setting | convertfrom-json).ticketingRequired
                                    }

                                    if ($value.ruleIdentifier -eq 'ApprovalRule') {
                                        $Activation['ApproveRequired'] = $($value.setting | convertfrom-json).Enabled
                                    }

                                }
                        }
                            # Eligible settings
                        if (($lifeCycleManagement.caller -eq 'Admin') -and ($lifeCycleManagement.level -eq 'Eligible')) {

                                foreach($value  in  $lifeCycleManagement.value) {
                                    if ($value.ruleIdentifier -eq 'ExpirationRule') {
                                        $Eligible['permanentAssignment'] = $($value.setting | convertfrom-json).permanentAssignment
                                        $Eligible['maximumGrantPeriodInMinutes'] = $($value.setting | convertfrom-json).maximumGrantPeriodInMinutes
                                    }

                                    if ($value.ruleIdentifier -eq 'MfaRule') {
                                        $Eligible['mfaRequired'] = $($value.setting | convertfrom-json).mfaRequired
                                    }
                                
                                }
                        }               
                        
                            # Active settings
                        if (($lifeCycleManagement.caller -eq 'Admin') -and ($lifeCycleManagement.level -eq 'Member')) {

                                    foreach($value  in  $lifeCycleManagement.value) {
                                        if ($value.ruleIdentifier -eq 'ExpirationRule') {
                                            $Active['permanentAssignment'] = $($value.setting | convertfrom-json).permanentAssignment
                                            $Active['maximumGrantPeriodInMinutes'] = $($value.setting | convertfrom-json).maximumGrantPeriodInMinutes
                                        }

                                        if ($value.ruleIdentifier -eq 'MfaRule') {
                                            $Active['mfaRequired'] = $($value.setting | convertfrom-json).mfaRequired
                                        }

                                        if ($value.ruleIdentifier -eq 'JustificationRule') {
                                            $Active['JustificationRequired'] = $($value.setting | convertfrom-json).required
                                        }

                                    }
                        }                    
                        

                    }


                    $rolesetting = [PSCustomObject]@{
                        "roleId" = $result.Id
                        "rolename" = $($result.roleDefinition.displayname)
                        "roledefinitionID" = $result.roleDefinitionId
                        "groupid" = $result.resourceId
                        "isDefault" = $result.isDefault
                        "activation"=  $activation
                        "Active"= $Active
                        "Eligible"=$Eligible
                    }
                    $rolesettings += $rolesetting
                }
            } 
            
            return $rolesettings
            
        } else {
            write-verbose "no roles found for PAG"
            return $NULL
        }

    }
}



# get PAG role assignments
function get-PIMGroupassignments
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(ParameterSetName='group', Mandatory=$True)]
        [String]$groupid,
        [Parameter(ParameterSetName='group',Mandatory=$false)]
        [String]$subjectID,
        [Parameter(ParameterSetName='group',Mandatory=$false)]
        [String]$roledefinitionID,
        [Parameter(ParameterSetName='group',Mandatory=$false)]
        [ValidateSet('Eligible','Active')]
        [String]$assignmentstate,
        [Parameter(Mandatory=$false)]
        [int]$count=100
    )
    Process
    {

        $API="privilegedAccess/aadGroups/roleAssignments"

        
       $querystring="`$expand=linkedEligibleRoleAssignment,subject,scopedResource,roleDefinition(`$expand=resource)"

            $queryString = "$querystring&`$filter=(roleDefinition/resource/id eq '$groupId')"
            # add filter for roledefinitionID, subjectID and assignmentState if exists
            if (![string]::IsNullOrEmpty($roledefinitionID)) {
                $queryString = "$querystring and (roledefinitionID eq '$roledefinitionID')"
            }

            if (![string]::IsNullOrEmpty($assignmentState)) {
                $queryString = "$querystring and (assignmentState eq '$assignmentstate')"
            }

            if (![string]::IsNullOrEmpty($subjectID)) {
                $queryString = "$querystring and (subjectId eq '$subjectID')"
            }

            # limit the query to return top 100 by default
            $queryString = "$querystring&`$count=true&`$orderby=roleDefinition/displayName&`$skip=0&`$top=$count"

        $results=Call-MSPIMAPI -AccessToken $AccessToken -API $API -QueryString $queryString
        
        if ($results) { 
            return $results | select @{N='assignmentid';E={$_.id}}, linkedEligibleRoleAssignmentid, @{N='groupid';E={$_.resourceid}} ,  roleDefinitionId,subjectId,memberType,startDateTime,endDateTime,assignmentState,status, subject
        } else {
            return $NULL
        }

    }
}


# delete assignment
function remove-PIMGroupassignments
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$groupid,
        [Parameter(Mandatory=$True)]
        [String]$roledefinitionID,
        [Parameter(Mandatory=$True)]
        [String]$subjectId,
        [Parameter(Mandatory=$false)]
        [String]$linkedEligibleRoleAssignmentId,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Eligible','Active')]
        [String]$assignmentstate='Eligible',
        [Parameter(Mandatory=$false)]
        [ValidateSet('AdminRemove','UserRemove')]
        [String]$type='AdminRemove',
        [Parameter(Mandatory=$false)]
        [String]$reason="Remove Assignment"
    )
    Process
    {

        $API="privilegedAccess/aadGroups/roleAssignmentRequests"

        if($assignmentstate -eq 'Active') {

            $assignment = get-PIMGroupassignments -groupid $groupid -roledefinitionID $roledefinitionID -subjectID $subjectId -assignmentstate $assignmentstate 
            if(![string]::IsNullOrEmpty($assignment.linkedEligibleRoleAssignmentId)){

                $body=@{
                    "linkedEligibleRoleAssignmentId"=$assignment.linkedEligibleRoleAssignmentId
                    "resourceId"=$groupID
                    "roleDefinitionId"=$roledefinitionID
                    "assignmentState"=$assignmentstate
                    "subjectId"=$subjectId
                    "scopedResourceId"=$null
                    "type"='UserRemove'
                    "reason"='Deactive assignment'
                } 
            } else {

                $body=@{
                    "resourceId"=$groupID
                   "roleDefinitionId"=$roledefinitionID
                   "assignmentState"=$assignmentstate
                   "subjectId"=$subjectId
                   "scopedResourceId"=$null
                   "type"='AdminRemove'
                   "reason"='remove assignment'
               } 
   
            }
           
        } else {

            $body=@{
                 "resourceId"=$groupID
                "roleDefinitionId"=$roledefinitionID
                "assignmentState"=$assignmentstate
                "subjectId"=$subjectId
                "scopedResourceId"=$null
                "type"='AdminRemove'
                "reason"='remove assignment'
            } 

        }

        Call-MSPIMAPI -AccessToken $AccessToken -API $API -method POST -body $body            
    
    }
}


# Add new assignment
function Add-PIMGroupassignments
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$groupid,
        [Parameter(Mandatory=$True)]
        [String]$roledefinitionID,
        [Parameter(Mandatory=$True)]
        [String]$subjectId,
        [Parameter(Mandatory=$false)]
        [String]$linkedEligibleRoleAssignmentId,
        [Parameter(Mandatory=$false)]
        [int]$startduration=0,  # default is null. if added, swtich the start time to new time
        [Parameter(Mandatory=$false)]
        [int]$duration=0,  # use time range like 1 day, 1 hour, default is 0 which will use maximum allowed time range
        [Parameter(Mandatory=$false)]
        [int]$minduration=10, # minumum duration to 10 minutes
        [Parameter(Mandatory=$false)]
        [ValidateSet('D','H','M')]
        [String]$durationunit='D', # D = day, H = hour, M=minutes.
        [Parameter(Mandatory=$false)]
        [String]$reason='Add userassignment to PAG',
        [Parameter(Mandatory=$false)]
        [ValidateSet('Eligible','Active')]
        [String]$assignmentstate='Eligible',
        [Parameter(Mandatory=$false)]
        [ValidateSet('AdminAdd','UserAdd')]
        [String]$type='AdminAdd'
    )
    Process
    {

        $API="privilegedAccess/aadGroups/roleAssignmentRequests"

        $assignments = get-PIMGroupassignments -groupid $groupid -assignmentstate $assignmentstate -subjectID $subjectId -roledefinitionID $roledefinitionID
        if ($assignments) {
            write-verbose "There is existing assignmens for subject: $subjectId. Will try update"            
            Update-PIMGroupassignments -groupid $groupId -roleDefinitionId $roleDefinitionId -assignmentstate $assignmentstate -subjectId $subjectID -duration $duration -startduration $startduration -durationunit $durationunit
            return $null
        }

        $role = get-PIMGrouprolesettings -groupid $groupid | where { $_.roledefinitionID -eq  $roledefinitionID}

        if (!$role) {
            write-error "invalid roledefinitionID "
            return $null
        }

        $rolesettings = $role.$($assignmentstate)

        # set time range
        Switch ($durationunit){
                "D" {$startdurationminutes = $startduration * 60 *24}
                "H" {$startdurationminutes = $startduration * 60}
                "M" {$startdurationminutes = $startduration}
                default {$startdurationminutes = $startduration * 60 *24}
        }
       

        $utcdate = $(get-date).ToUniversalTime() 
        $startutcdate = $(get-date -date $utcdate).AddMinutes($startdurationminutes)
        $starttime = $(get-date -date $startutcdate  -format "yyyy-MM-ddTHH:mm:ssZ").tostring()

        if($duration -eq 0) {
            # use no end time if no duraion giving and the role supports permanent
            if($rolesettings.permanentAssignment) {
                $endtime = $null
            } else {
                $endutcdate = $(get-date -date $startutcdate).AddMinutes($($rolesettings.maximumGrantPeriodInMinutes))
                $endtime =   $(get-date -date $endutcdate  -format "yyyy-MM-ddTHH:mm:ssZ").tostring()
           }
        
        } else {

            Switch ($durationunit){
                "D" {$durationminutes = $duration * 60 *24}
                "H" {$durationminutes = $duration * 60}
                "M" {$durationminutes = $duration}
                default {$durationminutes = $duration * 60 *24}
            }
            if ($durationminutes -gt $($rolesettings.maximumGrantPeriodInMinutes)) { 
                 $durationminutes = $($rolesettings.maximumGrantPeriodInMinutes)
            } elseif ($durationminutes -lt $minduration) { 
                $durationminutes = $minduration
           } 
           $endutcdate = $(get-date -date $startutcdate).AddMinutes($durationminutes)
           $endtime =   $(get-date -date $endutcdate  -format "yyyy-MM-ddTHH:mm:ssZ").tostring()

        }

        write-verbose "Will set assignment schedule time range: $starttime - $endtime"

        # set assignment parameters
        if ([string]::IsNullOrEmpty($linkedEligibleRoleAssignmentId)) {
            $body=@{
                "resourceId"=$groupID
                "roleDefinitionId"=$roledefinitionID
                "assignmentState"=$assignmentstate
                "subjectId"=$subjectId
                "type"=$type
                "reason"="add new member to PAG"
                "schedule" =@{
                    "type"="Once"
                    "startDateTime"=$starttime
                    "endDateTime"=$endtime
                }
            }  

        } else {
            $body=@{
                "linkedEligibleRoleAssignmentId"=$linkedEligibleRoleAssignmentId
                "resourceId"=$groupID
                "roleDefinitionId"=$roledefinitionID
                "assignmentState"=$assignmentstate
                "subjectId"=$subjectId
                "type"=$type
                "reason"=$reason
                "schedule" =@{
                    "type"="Once"
                    "startDateTime"=$starttime
                    "endDateTime"=$endtime
                }
            }  

        }
        

        
 
        $results=Call-MSPIMAPI -AccessToken $AccessToken -API $API -method POST -body $body            
        write-verbose $results

    }
}




# Update assignment
function Update-PIMGroupassignments
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$groupid,
        [Parameter(Mandatory=$True)]
        [String]$roledefinitionID,
        [Parameter(Mandatory=$True)]
        [String]$subjectId,
        [Parameter(Mandatory=$false)]
        [int]$startduration=0,  # default is null. if added, swtich the start time to new time
        [Parameter(Mandatory=$false)]
        [int]$duration=0,  # use time range like 1 day, 1 hour, default is 0 which will use maximum allowed time range
        [Parameter(Mandatory=$false)]
        [int]$minduration=10, # minumum duration to 10 minutes
        [Parameter(Mandatory=$false)]
        [ValidateSet('D','H','M')]
        [String]$durationunit='D', # D = day, H = hour, M=minutes.
        [Parameter(Mandatory=$false)]
        [String]$reason='Update Assignment Schedule',
        [Parameter(Mandatory=$false)]
        [ValidateSet('Eligible','Active')]
        [String]$assignmentstate='Eligible'
    )
    Process
    {

        $API="privilegedAccess/aadGroups/roleAssignmentRequests"
        $type='AdminUpdate'

        $assignment = get-PIMGroupassignments -groupid $groupid -assignmentstate $assignmentstate -subjectID $subjectId -roledefinitionID $roledefinitionID
        if (!$assignment) {
            write-verbose "no $assignmentstate assignment found for: $subjectId in group $groupid with role: $roledefinitionID; skip update"
            return $null
        }


        # for actived assignment, will remove the current activation and add a new activation assignment
        if(![string]::IsNullOrEmpty($assignment.linkedEligibleRoleAssignmentId -and $assignmentstate -eq 'Active')){
           
            write-verbose "there is existing activation assignment found for: $subjectId in group $groupid with role: $roledefinitionID; will do remove and add a new one"
            remove-PIMGroupassignments -groupid $groupid -subjectId $subjectId -roledefinitionID $roledefinitionID -assignmentstate 'Active'
            # add new activation window
            Activate-PIMGroupassignments -groupid $groupId -roleDefinitionId $roleDefinitionId -subjectId $subjectID -startduration $startduration -duration $duration -durationunit $durationunit
     
        }
       

        $role = get-PIMGrouprolesettings -groupid $groupid | where { $_.roledefinitionID -eq  $roledefinitionID}

        if (!$role) {
            write-error "invalid roledefinitionID "
            return $null
        }

        $rolesettings = $role.$($assignmentstate)

        # set time range
       Switch ($durationunit){
                 "D" {$startdurationminutes = $startduration * 60 *24}
                 "H" {$startdurationminutes = $startduration * 60}
                 "M" {$startdurationminutes = $startduration}
                 default {$startdurationminutes = $startduration * 60 *24}
             }
        
 
 
         $utcdate = $(get-date).ToUniversalTime() 
         $startutcdate = $(get-date -date $utcdate).AddMinutes($startdurationminutes)
         $starttime = $(get-date -date $startutcdate  -format "yyyy-MM-ddTHH:mm:ssZ").tostring()
 
         if($duration -eq 0) {
             # use no end time if no duraion giving and the role supports permanent
             if($rolesettings.permanentAssignment) {
                 $endtime = $null
             } else {
                 $endutcdate = $(get-date -date $startutcdate).AddMinutes($($rolesettings.maximumGrantPeriodInMinutes))
                 $endtime =   $(get-date -date $endutcdate  -format "yyyy-MM-ddTHH:mm:ssZ").tostring()
            }
         
         } else {
 
             Switch ($durationunit){
                 "D" {$durationminutes = $duration * 60 *24}
                 "H" {$durationminutes = $duration * 60}
                 "M" {$durationminutes = $duration}
                 default {$durationminutes = $duration * 60 *24}
             }
             if ($durationminutes -gt $($rolesettings.maximumGrantPeriodInMinutes)) { 
                  $durationminutes = $($rolesettings.maximumGrantPeriodInMinutes)
             } elseif ($durationminutes -lt $minduration) { 
                 $durationminutes = $minduration
            } 
            $endutcdate = $(get-date -date $startutcdate).AddMinutes($durationminutes)
            $endtime =   $(get-date -date $endutcdate  -format "yyyy-MM-ddTHH:mm:ssZ").tostring()
 
         }
 
         write-verbose "Will set assignment schedule time range: $starttime - $endtime"

        $body=@{
                "linkedEligibleRoleAssignmentId"=$($assignment.linkedEligibleRoleAssignmentId)
                "resourceId"=$groupID
                "roleDefinitionId"=$roledefinitionID
                "assignmentState"=$assignmentstate
                "subjectId"=$subjectId
                "type"=$type
                "reason"=$reason
                "schedule" =@{
                    "type"="Once"
                    "startDateTime"=$starttime
                    "endDateTime"=$endtime
                }
            }             
           
 
        $results=Call-MSPIMAPI -AccessToken $AccessToken -API $API -method POST -body $body            
        write-verbose $results

    }
}





# Activate Eligible assignment
function Activate-PIMGroupassignments
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$groupid,
        [Parameter(Mandatory=$True)]
        [String]$roledefinitionID,
        [Parameter(Mandatory=$True)]
        [String]$subjectId,
        [Parameter(Mandatory=$false)]
        [int]$startduration=0,  # default is 0. if added, swtich the start time to new time
        [Parameter(Mandatory=$false)]
        [int]$duration=0, # default is 0. if added, swtich the start time to new time
        [Parameter(Mandatory=$false)]
        [int]$minduration=30, 
        [Parameter(Mandatory=$false)]
        [ValidateSet('D','H','M')]
        [String]$durationunit='H', # D = day, H = hour, M=minutes.
        [Parameter(Mandatory=$false)]
        [String]$reason='Active Eligible Assignment'
    )
    Process
    {

        $assignmentstate='Eligible'

        $activeassignment = get-PIMGroupassignments -groupid $groupid -assignmentstate 'Active' -subjectID $subjectId -roledefinitionID $roledefinitionID
        if ($activeassignment) {
            write-verbose "There are existing Active assignment found for: $subjectId in group $groupid with role: $roledefinitionID; skip Activate"
            return $null
        }


        $assignment = get-PIMGroupassignments -groupid $groupid -assignmentstate $assignmentstate -subjectID $subjectId -roledefinitionID $roledefinitionID
        if (!$assignment) {
            write-verbose "no $assignmentstate assignment found for: $subjectId in group $groupid with role: $roledefinitionID; skip update"
            return $null
        }

       
        $role = get-PIMGrouprolesettings -groupid $groupid | where { $_.roledefinitionID -eq  $roledefinitionID}

        if (!$role) {
            write-error "invalid roledefinitionID "
            return $null
        }

        $activationsetting = $role.activation

        # set time range
        Switch ($durationunit){
            "D" {$startdurationminutes = $startduration * 60 *24}
            "H" {$startdurationminutes = $startduration * 60}
            "M" {$startdurationminutes = $startduration}
            default {$startdurationminutes = $startduration * 60 *24}
        }
        
    
        if($duration -eq 0) {
            # use max windows for activation assignment
            $durationminutes =  $activationsetting.ActivationDuration
                            
        } else {

            Switch ($durationunit){
                "D" {$durationminutes = $duration * 60 *24}
                "H" {$durationminutes = $duration * 60}
                "M" {$durationminutes = $duration}
                default {$durationminutes = $duration * 60 *24}
            }
            if ($durationminutes -gt $($activationsetting.ActivationDuration)) { 
                 $durationminutes = $activationsetting.ActivationDuration
            } elseif ($durationminutes -lt $minduration) { 
                $durationminutes = $minduration
           } 
          
        }

        write-verbose "Activation time range: $durationminutes minutes; do activation after $startdurationminutes minutes"
        # add linked assignment            
           
 
        Add-PIMGroupassignments -groupid $groupId -roleDefinitionId $roleDefinitionId -subjectId $subjectID -linkedEligibleRoleAssignmentId $($assignment.assignmentid) -startduration $startdurationminutes -duration $durationminutes -durationunit 'M' -assignmentstate 'Active' -type 'UserAdd' -reason "Active Eligible assignment"
          
    }
}





