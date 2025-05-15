# Begin Alma Portfolios link checking script
# Function to check if a URL is broken or redirected to a domain
function Test-Url {
    param (
        [string]$url
    )
    $maxRetries = 3
    $retryCount = 0
    $errorCode = $null

    while ($retryCount -lt $maxRetries -and $errorCode -eq $null) {
        try {
            $response = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 30 -Headers @{"User-Agent"="Mozilla/5.0"} -MaximumRedirection 5 -ErrorAction Stop
         	# Test for 400 Bad request and 404 Not Found
			if ($response.StatusCode -eq 400 -or $response.StatusCode -eq 404) {
                return $response.StatusCode
	# Test fo HTTP 500 error codes	
			elseif ($response.StatusCode -ge 500 -and $response.StatusCode -lt 600) {
                return "Server Error $($response.StatusCode)"
            }
	#Test if a long URL redirects to a domain
            } elseif ($response.StatusCode -eq 301 -or $response.StatusCode -eq 308) {
                $finalUrl = $response.Headers.Location
                if ($finalUrl -match "^https?://[^/]+/?$") {
                    return "$response.StatusCode - Redirected to domain"
                }
            }
        } catch {
            if ($_.Exception.Response.StatusCode -eq 404) {
                return $_.Exception.Response.StatusCode
	#Test for DNS errors, timeouts, expired SSL certificates and dead websites
            } elseif ($_.Exception -match "The remote name could not be resolved") {
                return "DNS Lookup Failed"
            } elseif ($_.Exception -match "The operation has timed out") {
                return "Timeout"
            } elseif ($_.Exception -match "The underlying connection was closed") {
                return "Connection Closed"
            } else {
                $errorCode = $null
            }
        }
        $retryCount++
        Start-Sleep -Seconds 5
    }
    return $errorCode
}

# Open the input Excel file
$inputFilename = Get-ChildItem -Path . -Filter "*_portfolios.xlsx" | Select-Object -First 1
$outputFilename = "broken-links.csv"

try {
    Write-Host "##################################################" -ForegroundColor DarkYellow
    Write-Host "Alma Portfolios link checking script (Version 1.0)" -ForegroundColor DarkYellow
    Write-Host "##################################################" -ForegroundColor DarkYellow
    Write-Host ""  # Blank line

    if ($inputFilename) {
        Write-Host "Checking $($inputFilename.Name)" -ForegroundColor Magenta
        Write-Host ""  # Blank line
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Open($inputFilename.FullName)
        $worksheet = $workbook.Sheets.Item(1)
        $range = $worksheet.UsedRange
        $output = @()

        $lineCount = 0
        for ($row = 2; $row -le $range.Rows.Count; $row++) {
            $lineCount++
            Write-Host "Processing line $lineCount"

            # Check URL in column BF
            $url = $range.Cells.Item($row, 58).Text
            $mmsId = $range.Cells.Item($row, 12).Text
            if ($url) {
                Write-Host "Checking URL: $url"
                $errorCode = Test-Url -url $url
                if ($errorCode -and $errorCode -ne $null) {
                    Write-Host "Broken link detected: $url - Status Code: $errorCode" -ForegroundColor Red
                    $output += [pscustomobject]@{
                        "MMS ID"         = $mmsId
                        "HTTP Error Code" = $errorCode
                    }
                } else {
                    Write-Host "URL OK: $url" -ForegroundColor Green
                }
                Write-Host ""  # Blank line
            }
        }

        # Export the results to a new CSV file
        $output | Export-Csv -Path $outputFilename -NoTypeInformation

        Write-Host "Link checking complete. Please open $outputFilename" -ForegroundColor Green
        $workbook.Close($false)
        $excel.Quit()
    } else {
        Write-Host "No Excel file found with the specified pattern."
    }
} catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}

# Keep the PowerShell window open
Read-Host -Prompt "Press Enter to exit"
# End Alma Portfilios link checking script
