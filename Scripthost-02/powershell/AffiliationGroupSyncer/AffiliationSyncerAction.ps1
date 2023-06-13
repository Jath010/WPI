Import-Module D:\wpi\powershell\AffiliationGroupSyncer\AffiliationGroupSyncer.ps1

Sync-AffiliationGroup -PrimaryAffiliation Student
Sync-AffiliationGroup -PrimaryAffiliation Student -SubAffiliation freshman
Sync-AffiliationGroup -PrimaryAffiliation Student -SubAffiliation sophomore
Sync-AffiliationGroup -PrimaryAffiliation Student -SubAffiliation junior
Sync-AffiliationGroup -PrimaryAffiliation Student -SubAffiliation senior
Sync-AffiliationGroup -PrimaryAffiliation Faculty
Sync-AffiliationGroup -PrimaryAffiliation Staff
Sync-AffiliationGroup -PrimaryAffiliation Affiliate