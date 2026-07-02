# Begin Alma Portfolios link checking script

function Test-Url {
    param (
        [string]$url
    )

    $maxRetries = 2
    $retryCount = 0
    $errorCode = $null

    while ($retryCount -lt $maxRetries -and $errorCode -eq $null) {
        try {
            # -MaximumRedirection 1 follows exactly one hop, which is enough to detect
            # whether a 301 has redirected to a bare domain, without silently resolving
            # the full chain (which would return a final 200 and hide the redirect).
            $response = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Head -TimeoutSec 90 `
                -Headers @{"User-Agent" = "Mozilla/5.0"} -MaximumRedirection 1 -ErrorAction Stop

            # ----- STEP 1: 301 redirected to a bare domain -----
            # No exception was thrown, so any redirect was followed and landed on a
            # working (2xx) page. Compare host+path to detect whether a redirect
            # occurred, then flag only if it landed on a bare top-level domain.
            $requestedUri = [System.Uri]$url
            $landedUri = $response.BaseResponse.ResponseUri
            $redirectOccurred = ($requestedUri.Host -ne $landedUri.Host) -or
                                 ($requestedUri.AbsolutePath.TrimEnd('/') -ne $landedUri.AbsolutePath.TrimEnd('/'))

            if ($redirectOccurred -and $landedUri.AbsoluteUri -match "^https?://[^/]+/?$") {
                return "301 - Redirected to domain"
            }

            # ----- STEP 2: 404, 405, 500 -----
            if ($response.StatusCode -eq 404) {
                return $response.StatusCode
            } elseif ($response.StatusCode -eq 405) {
                return $response.StatusCode
            } elseif ($response.StatusCode -eq 500) {
                return "Server Error $($response.StatusCode)"
            }

        } catch {
            $statusCode = $null
            if ($_.Exception.Response -ne $null) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            # ----- STEP 1: 301 redirected to a bare domain -----
            # A thrown 3xx here means a second redirect occurred beyond the one hop
            # we allowed; PowerShell surfaces that as an exception.
            if ($statusCode -eq 301) {
                $locationHeader = $_.Exception.Response.Headers["Location"]
                $finalUrl = [string]$locationHeader

                if ($finalUrl -match "^https?://[^/]+/?$") {
                    return "301 - Redirected to domain"
                }
                # 301 to a specific path, not a bare domain - not flagged.
            }
            # ----- STEP 2: 404, 405, 500 -----
            elseif ($statusCode -eq 404) {
                return $statusCode
            } elseif ($statusCode -eq 405) {
                return $statusCode
            } elseif ($statusCode -eq 500) {
                return "Server Error $statusCode"
            }
            # ----- STEP 2: named connection/DNS errors -----
            elseif ($_.Exception -match "The remote name could not be resolved") {
                return "DNS Lookup Failed"
            } elseif ($_.Exception -match "The operation has timed out") {
                return "Timeout"
            } elseif ($_.Exception -match "The underlying connection was closed") {
                return "Connection Closed"
            } elseif ($_.Exception -match "NXDOMAIN") {
                return "NXDOMAIN Error"
            } else {
                $errorCode = $null
            }
        }

        $retryCount++
        Start-Sleep -Seconds 1
    }

    return $errorCode
}

# Open the input Excel file
$inputFilename = Get-ChildItem -Path . -Filter "*_portfolios.xlsx" | Select-Object -First 1

$outputFilename = "broken-links.csv"

try {
    Write-Host "##################################################" -ForegroundColor DarkYellow
    Write-Host "Alma Portfolios link checking script (Version 1.2)" -ForegroundColor DarkYellow
    Write-Host "##################################################" -ForegroundColor DarkYellow
    Write-Host "" # Blank line

    if ($inputFilename) {
        Write-Host "Checking $($inputFilename.Name)" -ForegroundColor Magenta
        Write-Host "" # Blank line

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
                        "MMS ID"          = $mmsId
                        "HTTP Error Code" = $errorCode
                    }
                } else {
                    Write-Host "URL OK: $url" -ForegroundColor Green
                }

                Write-Host "" # Blank line
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

# End Alma Portfolios link checking script
