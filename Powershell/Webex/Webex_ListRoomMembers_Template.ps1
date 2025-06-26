# Remember, if you want to manipulate a room, YOU must be an admin for that room.
# By default, this script will output results into two .txt files on your desktop.
 
 # Access token
$accessToken = "ACCESS_TOKEN" # <-- replace with your Access Token

# Define named rooms
$rooms = @(
    @{ Name = "ROOM_NAME"; RoomId = "ROOM_ID" }, # <--- 
    @{ Name = "ROOM_NAME"; RoomId = "ROOM_ID" }, # <-- replace with desired Room Name and Room ID
    @{ Name = "ROOM_NAME"; RoomId = "ROOM_ID" }, # <--- If you do not have the room IDs, you can use my Webex_GetRoomIDs.ps1 script.
    # Add more if needed
)

# Headers
$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

# Output file paths
$desktopPath = [Environment]::GetFolderPath("Desktop") # <-- Change file output location if desired
$timestamp = Get-Date -Format "MM-dd-yyyy" # <-- Change to whatever date/time format you want. Default is American format.
$fullOutputFile = Join-Path $desktopPath "WebexMembers_ADCheck_$timestamp.txt" # <-- Outputs the full list of users in the room. Change to determine output file name and type.
$missingOutputFile = Join-Path $desktopPath "WebexMembers_NotFoundInAD_$timestamp.txt" # <-- Outputs just the users missing from AD. Change to determine output file name and type.

# Output arrays
$fullOutput = @("Members in Webex Rooms with AD Verification:")
$missingOutput = @("Members NOT Found in Active Directory:")

# Loop through each room
foreach ($room in $rooms) {
    $roomName = $room.Name
    $roomId = $room.RoomId
    $url = "https://webexapis.com/v1/memberships?roomId=$roomId" # <-- The main API query that makes this work

    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get

    $fullOutput += "`nRoom: $roomName"
    $missingOutput += "`nRoom: $roomName"

    if ($response.items) {
        $matchingMembers = $response.items | Where-Object { $_.personEmail -like "*EMAIL" } # <-- replace EMAIL with email you wish to filter for, make sure to leave the * symbol.

        if ($matchingMembers.Count -eq 0) {
            $fullOutput += "- No EMAIL members found." # <-- replace EMAIL with the email you wish to filter for.
        } else {
            foreach ($member in $matchingMembers) {
                $email = $member.personEmail
                $name = $member.personDisplayName
                                            # add -searchbase here to specify a path in AD, or leave it out to search the whole domain.
                # AD lookup                                     V
                $adUser = Get-ADUser -Filter { mail -eq $email } -Properties mail, Enabled -ErrorAction SilentlyContinue

                if ($adUser) {
                    if ($adUser.Enabled) {
                        $fullOutput += "- $name ($email) Found in AD (Enabled)"
                    } else {
                        $line = "- $name ($email) Found in AD (Disabled)"
                        $fullOutput += $line
                        $missingOutput += $line
                    }
                } else {
                    $line = "- $name ($email) Not found in AD"
                    $fullOutput += $line
                    $missingOutput += $line
                }
            }
        }
    } else {
        $fullOutput += "- Failed to retrieve members or room is empty."
        $missingOutput += "- Failed to retrieve members or room is empty."
    }
}

# Write both output files
$fullOutput | Tee-Object -FilePath $fullOutputFile | ForEach-Object { Write-Host $_ }
$missingOutput | Set-Content -Path $missingOutputFile

Write-Host "`nSaved full report to: $fullOutputFile"
Write-Host "Saved missing AD entries to: $missingOutputFile"
