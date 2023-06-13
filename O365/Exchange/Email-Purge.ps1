Connect-IPPSSession

New-ComplianceSearch -name testsearch3 -ExchangeLocation all -ContentMatchQuery ‘subject:”Test Com search 1” AND attachmentnames:attachment1.txt‘

Start-ComplianceSearch testsearch3

get-ComplianceSearch testsearch3 |fl

New-ComplianceSearchAction -SearchName "Query - SGAMistake" -Purge -PurgeType softdelete