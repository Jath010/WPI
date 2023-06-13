New-GlobalAddressList -Name "ABPTest_GAL" -IncludedRecipients AllRecipients -ConditionalCustomAttribute15 "ABPTest"
New-OfflineAddressBook -Name "OAB_ABPTest" -AddressLists "\ABPTest_GAL"
New-AddressList -Name "ABPTest Rooms" -RecipientFilter "Alias -eq 'NotARoom' -and (RecipientDisplayType -eq 'ConferenceRoomMailbox' -or RecipientDisplayType -eq 'SyncedConferenceRoomMailbox')"
New-AddressList -Name "All ABPTest" -IncludedRecipients MailboxUsers -ConditionalCustomAttribute15 "ABPTest"
New-AddressBookPolicy -Name "ABPTest" -GlobalAddressList "\ABPTest_GAL" -OfflineAddressBook "OAB_ABPTest" -AddressLists "\All ABPTest" -RoomList "\ABPTest Rooms"