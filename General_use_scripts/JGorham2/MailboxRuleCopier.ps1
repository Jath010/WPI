function Copy-InboxRule {
    [CmdletBinding()]
    param (
        $SourceMailbox,
        $TargetMailbox
    )
    
    begin {
        $Rules = Get-InboxRule $Mailbox
        $params = (get-command set-inboxrule).parameters
    }
    
    process {
        foreach ($Rule in $Rules) {
            <# $Rule is the current item #>
            New-InboxRule -Name $rule.name -Mailbox $TargetMailbox
            foreach ($property in $Rule.PSObject.Properties) {
                <# $property is the current item #>
                #splatting is the answer
                if ($params.ContainsKey($property.name)) {
                    $command = @{ $property.name = $property.value }
                    Set-InboxRule -Mailbox $TargetMailbox -Identity $rule.name @command
                }
            }
        }
    }
    
    end {
        
    }
}