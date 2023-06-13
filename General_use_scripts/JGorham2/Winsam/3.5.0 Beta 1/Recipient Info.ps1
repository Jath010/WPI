Clear-Host
 <# 



     #RecipientTypes
        * UserMailbox - Includes UserMailbox, SharedMailbox, RoomMailbox; all mailboxes
        * MailUniversalDistributionGroup - Includes Unified Groups and all Exchange Distribution Groups (both online and on-prem)
            * GroupMailbox = UnifiedGroup
            * MailUniversalDistributionGroup = Distribution Groups (on-prem and online)
                * On-prem will have Capabilities = {MasteredOnPremise}
        * MailUniversalSecurityGroup - On-prem mail enabled security groups
                * On-prem will have Capabilities = {MasteredOnPremise}
#>
 
 
<# 
$alias = 'tcollins' 
$recip_user = Get-Recipient 'tcollins@wpi.edu' 
$Recip_resource_mbx = Get-Recipient 'its@wpi.edu'
$recip_resource_room = Get-Recipient 'calendar_ccc_conf@wpi.edu'
$recip_grp = Get-Recipient 'iamcore@wpi.edu'
$recip_local_distro = Get-Recipient 'sysops@wpi.edu'
$recip_online_distro = Get-Recipient 'dl-potpourri@wpi.edu'

$recip = @()
$recip += Get-Recipient 'tcollins@wpi.edu' 
$Recip += Get-Recipient 'its@wpi.edu'
$recip += Get-Recipient 'calendar_ccc_conf@wpi.edu'
$recip += Get-Recipient 'iamcore@wpi.edu'
$recip += Get-Recipient 'sysops@wpi.edu'
$recip += Get-Recipient 'dl-potpourri@wpi.edu'
$recip += Get-Recipient 'accounts@wpi.edu'
#>
<#
Name                                            RecipientType                  RecipientTypeDetails          
----                                            -------------                  --------------------          
Collins, Thomas L (73674)                       UserMailbox                    UserMailbox                   
IT Services                                     UserMailbox                    SharedMailbox                 
IT Conference Room                              UserMailbox                    RoomMailbox                   
gr-iamcore_292d8b62-8fb6-4cc5-b18b-d5f67b140949 MailUniversalDistributionGroup GroupMailbox                  
System_Operations_Team                          MailUniversalSecurityGroup     MailUniversalSecurityGroup    
dl-potpourri                                    MailUniversalDistributionGroup MailUniversalDistributionGroup
axp-advisor                                     MailContact                    MailContact


Count RecipientType
----- ----
35757 MailUser: GuestMailUser, MailUser
19699 UserMailbox: UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox, DiscoveryMailbox
 7878 MailUniversalDistributionGroup: GroupMailbox, MailUniversalDistributionGroup, RoomList
   87 MailContact: MailContact
   60 MailUniversalSecurityGroup: MailUniversalSecurityGroup
    7 DynamicDistributionGroup: DynamicDistributionGroup
#>


$alias = 'sysops@wpi.edu'

if ($alias -notmatch '@' -and $alias -notmatch 'wpi.edu') {Write-Host "Please enter a valid WPI email address";break}

$RecipientInfo = Get-Recipient $alias -ErrorAction SilentlyContinue


switch ($RecipientInfo.RecipientType) {
    'UserMailbox' {write-host 'Launch Mailbox View'}
    'MailUniversalDistributionGroup' {
        if ($RecipientInfo.RecipientTypeDetails -eq 'GroupMailbox') {write-host 'Launch Unified Group View'}
        if ($RecipientInfo.RecipientTypeDetails -eq 'MailUniversalDistributionGroup') {write-host 'Launch DistributionGroup View'}
        }
    'MailUniversalSecurityGroup' {write-host 'Launch DistributionGroup View'}
    default {Write-Host 'default'}
    }




