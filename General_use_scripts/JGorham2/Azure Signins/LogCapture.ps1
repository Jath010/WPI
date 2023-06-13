# Geoloc blocking signin audit script
# Joshua Gorham 5/10/2022
#

if (!((Get-Module).name -contains "azureadpreview")) {
    if (!((Get-installedModule).name -contains "azureadpreview")) {
        Install-Module azureadpreview
    }
    import-module azureadpreview -force
}

#$country = read-host "Please enter Country Code to search for:"
#$filter = "appID ne '52426cb4-e423-4179-8fb6-4d15c0cfb472' and appID ne '4f7fddca-a340-425f-9374-385bb2b74494' and appID ne '00000002-0000-0ff1-ce00-000000000000' and AppliedConditionalAccessPolicies/any(SignInAuditLogObjectAppliedConditionalAccessPolicies: SignInAuditLogObjectAppliedConditionalAccessPolicies/ID eq '9f0aff81-ee7e-47d8-91ac-5922e6717982' and location/countryOrRegion eq '$($Country)')"
$filter = "appID ne '52426cb4-e423-4179-8fb6-4d15c0cfb472' and appID ne '4f7fddca-a340-425f-9374-385bb2b74494' and appID ne '00000002-0000-0ff1-ce00-000000000000' and AppliedConditionalAccessPolicies/any(SignInAuditLogObjectAppliedConditionalAccessPolicies: SignInAuditLogObjectAppliedConditionalAccessPolicies/ID eq '9f0aff81-ee7e-47d8-91ac-5922e6717982')"

$hits = Get-AzureADAuditSignInLogs -Filter $filter

#requested output: Username and application, going to also add in time
$hits | select-object UserDisplayName, UserPrincipalName, @{N=’RequestID’; E={$_.Id}}, AppDisplayName, @{N=’Country’; E={$_.Location.CountryOrRegion}}, CreatedDateTime |Export-Csv -Path "C:\OFACSigninAudit.csv" -NoTypeInformation

<#

# Slate AppID       : 52426cb4-e423-4179-8fb6-4d15c0cfb472
# Workday AppID     : 4f7fddca-a340-425f-9374-385bb2b74494
# Exchange AppID    : 00000002-0000-0ff1-ce00-000000000000

# filter for time: CreatedDateTime gt 2022-05-10

#Get-AzureADAuditSignInLogs -Filter "appId ne '52426cb4-e423-4179-8fb6-4d15c0cfb472'"
#Get-AzureADAuditSignInLogs -Filter "location/city eq 'Redmond' and location/state eq 'Washington' and location/countryOrRegion eq 'US'"


{class SignInAuditLogObjectAppliedConditionalAccessPolicies {
                                     Id: 2f8713a2-646a-4b1d-a6f5-3f298406c199
                                     DisplayName: MFA Enabled Off Campus
                                     EnforcedGrantControls: class EnforcedGrantControls {
                                   }



Id                                   DisplayName                                  EnforcedGrantControls Result
--                                   -----------                                  --------------------- ------
2f8713a2-646a-4b1d-a6f5-3f298406c199 MFA Enabled Off Campus                       {Mfa}                 success
6b0e0bb3-65e6-423c-a7c8-d4d0b844da3a Workday MFA Policy (w/ Training Exclusion)   {}                    notEnabled
d0611471-0196-4a44-9ba7-fd912df6992d MFA App Policy (Non Production)              {Mfa}                 notApplied
0babbf9a-a8ae-4a2d-92b8-33945bc78703 Canvas Beta MFA Policy                       {}                    notEnabled
f42ef65a-719a-4d05-97f6-4c59c05be126 Banned IP Block                              {Block}               notApplied
bccce084-acf7-437c-b073-3ecfbc21563c Test - SalesForce-partialSB-CAAC             {}                    notEnabled
385735bf-a5b5-4c69-9216-0702e718e437 MFA App Policy (Production)                  {Mfa}                 notApplied
842b4dc6-0d60-4d7a-9517-84485a59a8a4 Workday Preview CAAC                         {}                    notEnabled
3ded1d44-0630-4a0d-b01a-efaa5aa7864f MFA-All Internal/External                    {Mfa}                 notApplied
9c0e5e38-6beb-4cda-bd2c-803ce97da36d Global Admin MFA                             {Mfa}                 notApplied
8e37a2ee-dd85-41ea-920f-1aa958747f0e Chemistry Stockroom Restriction              {Block}               notApplied
ab288788-d845-408c-a863-017b2120d134 lleclerc - test AAD view restriction         {}                    notEnabled
c99174d4-d57d-46b4-a0a4-29d080ab9994 Alum - Block portal.azure.com access         {Block}               notApplied
538bfcfa-64c6-4337-9808-54cb9f9a9400 Block users that request no email access     {Block}               notApplied
e889456b-4e52-4f65-b872-6a07df70a589 Enforce MFA on Tableau Admins on Tableau App {Mfa}                 reportOn...
9f0aff81-ee7e-47d8-91ac-5922e6717982 OFAC Geoblock                                {Block}               reportOn...


Get-AzureADAuditSignInLogs -Filter "AppliedConditionalAccessPolicies/any(class: SignInAuditLogObjectAppliedConditionalAccessPolicies/DisplayName eq 'OFAC Geoblock')"
Get-AzureADAuditSignInLogs -Filter "AppliedConditionalAccessPolicies/any(class: SignInAuditLogObjectAppliedConditionalAccessPolicies/ID eq '9f0aff81-ee7e-47d8-91ac-5922e6717982')"
Get-AzureADAuditSignInLogs -Filter "AppliedConditionalAccessPolicies/any(SignInAuditLogObjectAppliedConditionalAccessPolicies: SignInAuditLogObjectAppliedConditionalAccessPolicies/ID eq '9f0aff81-ee7e-47d8-91ac-5922e6717982')"
#>