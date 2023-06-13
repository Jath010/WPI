#Create a RoomList

function New-RoomList {
    param (
        #Name needs to be a string
        $RoomName
    )
    $email = "RoomList_$($RoomName.Replace(" ","_"))@wpi.edu"
    New-DistributionGroup -Name $RoomName -PrimarySmtpAddress $email -RoomList
}

function Set-RoomPlaceData {
    param (
        $RoomAddress
    )
    Set-Mailbox $RoomAddress -CountryOrRegion US -State "Massachusetts" -City "Worcester" #-Floor 1 -FloorLabel "Ground" -Capacity 5
}

function Get-RoomLists {
    Get-DistributionGroup -Filter "PrimarySmtpAddress -like 'RoomList_*'"
}