# This script will pull the room IDs of ALL the rooms YOU are in. So if you want to manipulate a room, YOU must also be in the room, AND an admin for the room.
# By default, it will output results into a .txt file on your desktop.

 $accessToken = "ACCESS_TOKEN" # <-- Replace with your Webex Developer access token
$url = "https://webexapis.com/v1/rooms"

$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

$response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get

# Output file path on desktop
$desktopPath = [Environment]::GetFolderPath("Desktop") # <-- Change to determine where to have the output file saved
$OutputPath = Join-Path $desktopPath "Webex_RoomIds$timestamp.txt" # <-- Change to determine what format the output file will be

# Clear or create the file first
"" | Out-File -FilePath $outputPath

foreach ($room in $response.items) {
    $output = @()
    $output += "Room Title: $($room.title)"
    $output += "Room ID: $($room.id)"
    $output += ""
    $output | Out-File -FilePath $outputPath -Append
}
