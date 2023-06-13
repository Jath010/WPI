function Sync-AADGroups {
    [CmdletBinding()]
    param (
        $TargetGroupID,
        $ReferenceGroupID
    )
    
    begin {
        $TargetMembers = Get-AzureADGroupMember -ObjectId $TargetGroupID -All:$true
        $ReferenceMembers = Get-AzureADGroupMember -ObjectId $ReferenceGroupID -All:$true
    }
    
    process {
        if ($null -eq $TargetMembers) {
            $AddMembers = $ReferenceMembers
        }
        elseif ($null -eq $ReferenceMembers) {
            $RemoveMembers = $TargetMembers
        }
        else {
            #reconscile lists
            $comparisons = Compare-Object  $TargetMembers $ReferenceMembers -Property ObjectId
                
            # Store the users who should and shouldn't be in the lists in variables.
            $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object ObjectId
            $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object ObjectId
        }
    }
    
    end {
        # Iterate through the add/remove lists and do what's necessary.
        ForEach ($Removal in $RemoveMembers.ObjectId) {
            Remove-AzureADGroupMember -ObjectId $TargetGroupID -MemberId $Removal
        }
                        
        ForEach ($Addition in $AddMembers.ObjectId) {
            Add-AzureADGroupMember -ObjectId $TargetGroupID -RefObjectId $Addition
                        
        }
    }
}