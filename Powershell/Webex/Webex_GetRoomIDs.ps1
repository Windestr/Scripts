$accessToken = "ACCESS_TOKEN"
$url = "https://webexapis.com/v1/rooms"

$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

$response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get

# Output file path on desktop
$desktopPath = [Environment]::GetFolderPath("Desktop")
$OutputPath = Join-Path $desktopPath "Webex_RoomIds$timestamp.txt"

# Clear or create the file first
"" | Out-File -FilePath $outputPath

foreach ($room in $response.items) {
    $output = @()
    $output += "Room Title: $($room.title)"
    $output += "Room ID: $($room.id)"
    $output += ""
    $output | Out-File -FilePath $outputPath -Append
}
