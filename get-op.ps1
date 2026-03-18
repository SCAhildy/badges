$baseUrl = "https://trimarisop.azurewebsites.net"
$startNumber = 1
$endNumber = 3500

Write-Output "Starting bulk download for records $startNumber through $endNumber..."

for ($i = $startNumber; $i -le $endNumber; $i++) {
    
    # Format the number with leading zeros (e.g., 1 becomes "000001")
    $targetNumber = "{0:D6}" -f $i
    $targetUrl = "$baseUrl/op?q=$targetNumber"
    
    Write-Output "Processing record $targetNumber..."
    
    # 1. Fetch the webpage content
    try {
        # -ErrorAction Stop ensures that 404 pages or connection errors get caught
        $response = Invoke-WebRequest -Uri $targetUrl -UseBasicParsing -ErrorAction Stop
        $html = $response.Content
    } catch {
        Write-Warning " -> Failed to fetch $targetNumber (Page might not exist). Skipping."
        continue # Skip the rest of this loop and move to the next number
    }

    # 2. Extract the Person's Name
    $nameRegex = '(?i)<h1[^>]*text-align:center[^>]*>(.*?)</h1>'
    $personName = ""

    if ($html -match $nameRegex) {
        $personName = $matches[1].Trim()
    }

    # 3. Extract and Download the Image
    $imageRegex = "(?i)src=['""]([^'""]*$targetNumber[^'""]*)['""]"

    if ($html -match $imageRegex) {
        $imagePath = $matches[1]
        $imageUrl = $baseUrl + $imagePath
        $extension = [System.IO.Path]::GetExtension($imagePath)
        
        # Create safe filename
        if ($personName) {
            $safeName = $personName -replace '[<>:"/\\|?*]', ''
            $fileName = "$safeName$extension"
        } else {
            $fileName = Split-Path $imagePath -Leaf
        }
        
        $downloadPath = Join-Path -Path $PWD -ChildPath $fileName
        
        try {
            Invoke-WebRequest -Uri $imageUrl -OutFile $downloadPath -UseBasicParsing -ErrorAction Stop
            Write-Output " -> Successfully downloaded: $fileName"
        } catch {
            Write-Warning " -> Failed to download the image for $targetNumber."
        }

    } else {
        Write-Output " -> No image found on page for $targetNumber."
    }

    # 4. Be polite to the server: Pause for 1000 milliseconds
    Start-Sleep -Milliseconds 1000
}

Write-Output "Bulk download complete!"