Import-Module \\scripthost-02\D$\wpi\powershell\AffiliationGroupSyncer\AffiliationGroupSyncer.ps1

Sync-AffiliationGroup -PrimaryAffiliation Student
Sync-AffiliationGroup -PrimaryAffiliation Faculty
Sync-AffiliationGroup -PrimaryAffiliation Staff
Sync-AffiliationGroup -PrimaryAffiliation Affiliate