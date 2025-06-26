# Access token
$accessToken = "ACCESS_TOKEN" # <-- replace with your Access Token

# Define named rooms
$rooms = @(
    @{ Name = "ROOM_NAME"; RoomId = "ROOM_ID" }, # <--- 
    @{ Name = "ROOM_NAME"; RoomId = "ROOM_ID" }, # <-- replace with desired Room Name and Room ID
    @{ Name = "ROOM_NAME"; RoomId = "ROOM_ID" }, # <---
    # Add more if needed
)

# Headers
$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

# Output file paths
$desktopPath = [Environment]::GetFolderPath("Desktop") # <-- Change file output location if desired
$timestamp = Get-Date -Format "MM-dd-yyyy"
$fullOutputFile = Join-Path $desktopPath "WebexMembers_ADCheck_$timestamp.txt" # <-- Outputs the full list
$missingOutputFile = Join-Path $desktopPath "WebexMembers_NotFoundInAD_$timestamp.txt" # <-- Outputs just the users missing from AD

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
            $fullOutput += "- No EMAIL members found." # <-- replace EMAIL with email you wish to filter for
        } else {
            foreach ($member in $matchingMembers) {
                $email = $member.personEmail
                $name = $member.personDisplayName
                                            # add -searchbase here to specify a path in AD, or leave it out to search the whole domain
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