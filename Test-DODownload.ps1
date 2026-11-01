<#
.SYNOPSIS
    Tests Invoke-DODownload.ps1 by performing:
    1. Full file download (verified via SHA-256)
    2. Partial byte-range download (verified byte-for-byte)
    3. Custom HTTP headers & HTTP Authentication (verified via a local HTTP server)
    4. Mid-download Cancellation / Abort & File Unlock/Removal verification
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$DownloaderScript = Join-Path $ScriptDir "Invoke-DODownload.ps1"
$OutputDir = Join-Path $ScriptDir "test-output"

if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

$TestUrl = "https://raw.githubusercontent.com/PowerShell/PowerShell/master/README.md"
$TargetFullFile    = Join-Path $OutputDir "PowerShell-README.md"
$ReferenceFile     = Join-Path $OutputDir "Reference-README.md"
$TargetRangeFile   = Join-Path $OutputDir "PowerShell-README-Range512.bin"
$TargetAuthFile    = Join-Path $OutputDir "Authenticated-Secret.txt"
$TargetCancelFile  = Join-Path $OutputDir "Canceled-Download.bin"

foreach ($f in @($TargetFullFile, $ReferenceFile, $TargetRangeFile, $TargetAuthFile)) {
    if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
}

Write-Host "=== Test 1: Full File Download via Delivery Optimization ===" -ForegroundColor Yellow
& $DownloaderScript -Uri $TestUrl -OutFile $TargetFullFile -DisplayName "DO Full File Test"
$doFile = Get-Item -LiteralPath $TargetFullFile

# Download reference copy via Invoke-WebRequest to compare SHA-256
Invoke-WebRequest -Uri $TestUrl -OutFile $ReferenceFile -UseBasicParsing
$doHash  = (Get-FileHash -LiteralPath $TargetFullFile -Algorithm SHA256).Hash
$refHash = (Get-FileHash -LiteralPath $ReferenceFile  -Algorithm SHA256).Hash

if ($doHash -ne $refHash) {
    throw "Test 1 FAILED: SHA-256 mismatch! DO=$doHash vs Ref=$refHash"
}
Write-Host "[PASS] Full file download verified! Size: $($doFile.Length) bytes | SHA-256: $doHash`n" -ForegroundColor Green

Write-Host "=== Test 2: Partial Byte-Range Download (First 512 bytes) ===" -ForegroundColor Yellow
& $DownloaderScript -Uri $TestUrl -OutFile $TargetRangeFile -DisplayName "DO Range Test" -RangeOffset 0 -RangeLength 512
$rangeFile = Get-Item -LiteralPath $TargetRangeFile

if ($rangeFile.Length -ne 512) {
    throw "Test 2 FAILED: Expected 512 bytes, got $($rangeFile.Length) bytes."
}

$fullBytes  = [System.IO.File]::ReadAllBytes($ReferenceFile)[0..511]
$rangeBytes = [System.IO.File]::ReadAllBytes($TargetRangeFile)
$diff = Compare-Object -ReferenceObject $fullBytes -DifferenceObject $rangeBytes -SyncWindow 0
if ($null -ne $diff) {
    throw "Test 2 FAILED: First 512 bytes do not match reference file."
}
Write-Host "[PASS] Partial range download (512 bytes) verified byte-for-byte!`n" -ForegroundColor Green

Write-Host "=== Test 3: HTTP Authentication & Custom Headers Verification (Local Auth Server) ===" -ForegroundColor Yellow
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
$listener.Start()
$port = ($listener.LocalEndpoint).Port
$expectedToken = "my-secret-token-999"
$expectedCustom = "PowerShell-DO-Client"
$secretText = "Authenticated secret payload delivered via Windows Delivery Optimization! " * 20
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($secretText)

$job = Start-Job -ScriptBlock {
    param($scriptPath, $port, $outFile, $token, $customVal)
    & $scriptPath `
        -Uri "http://127.0.0.1:$port/secret.txt" `
        -OutFile $outFile `
        -Headers @{
            "Authorization"   = "Bearer $token"
            "X-Custom-Client" = $customVal
        } `
        -DisplayName "DO Auth Test"
} -ArgumentList $DownloaderScript, $port, $TargetAuthFile, $expectedToken, $expectedCustom

$receivedAuthHeader = $false
$receivedCustomHeader = $false

try {
    while ($job.State -eq 'Running') {
        if ($listener.Pending()) {
            $client = $listener.AcceptTcpClient()
            $stream = $client.GetStream()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII, $false, 1024, $true)
            $reqLines = @()
            while ($null -ne ($line = $reader.ReadLine()) -and $line -ne '') {
                $reqLines += $line
            }

            $hasAuth   = [bool]($reqLines | Where-Object { $_ -eq "Authorization: Bearer $expectedToken" })
            $hasCustom = [bool]($reqLines | Where-Object { $_ -eq "X-Custom-Client: $expectedCustom" })

            if ($hasAuth -and $hasCustom) {
                $receivedAuthHeader = $true
                $receivedCustomHeader = $true
                $len = $bodyBytes.Length
                $respHeader = "HTTP/1.1 200 OK`r`nContent-Type: text/plain`r`nAccept-Ranges: bytes`r`nContent-Length: $len`r`nConnection: close`r`n`r`n"
                $respBytes = [System.Text.Encoding]::ASCII.GetBytes($respHeader)
                $stream.Write($respBytes, 0, $respBytes.Length)
                if ($reqLines[0] -notmatch '^HEAD ') {
                    $stream.Write($bodyBytes, 0, $bodyBytes.Length)
                }
            } else {
                $respHeader = "HTTP/1.1 401 Unauthorized`r`nWWW-Authenticate: Bearer`r`nContent-Length: 0`r`nConnection: close`r`n`r`n"
                $respBytes = [System.Text.Encoding]::ASCII.GetBytes($respHeader)
                $stream.Write($respBytes, 0, $respBytes.Length)
            }
            $stream.Flush()
            $client.Close()
        }
        Start-Sleep -Milliseconds 50
    }

    Receive-Job -Job $job -Wait -ErrorAction Stop | Out-Null
}
finally {
    $listener.Stop()
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
}

if (-not $receivedAuthHeader -or -not $receivedCustomHeader) {
    throw "Test 3 FAILED: Server did not receive expected Authorization or X-Custom-Client headers!"
}

$downloadedText = [System.IO.File]::ReadAllText($TargetAuthFile, [System.Text.Encoding]::UTF8)
if ($downloadedText -ne $secretText) {
    throw "Test 3 FAILED: Downloaded authenticated file content does not match expected payload!"
}
Write-Host "[PASS] HTTP Auth ('Authorization: Bearer ...') & Custom Header verified by server and downloaded ($($bodyBytes.Length) bytes)!`n" -ForegroundColor Green

Write-Host "=== Test 4: Cancel / Abort & Temp File Unlock Verification ===" -ForegroundColor Yellow
$CancelDir = Join-Path $OutputDir "cancel-test"
if (Test-Path $CancelDir) {
    Remove-Item $CancelDir -Recurse -Force
}
New-Item -Path $CancelDir -ItemType Directory -Force | Out-Null
$TargetCancelFile = Join-Path $CancelDir "Canceled-Download.bin"

$slowListener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
$slowListener.Start()
$slowPort = ($slowListener.LocalEndpoint).Port
$openClients = New-Object System.Collections.Generic.List[System.Net.Sockets.TcpClient]

$cancelJob = Start-Job -ScriptBlock {
    param($scriptPath, $port, $outFile)
    & $scriptPath -Uri "http://127.0.0.1:$port/slow.bin" -OutFile $outFile -DisplayName "DO Slow Cancel Test"
} -ArgumentList $DownloaderScript, $slowPort, $TargetCancelFile

$transferStarted = $false

try {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt 10) {
        while ($slowListener.Pending()) {
            $c = $slowListener.AcceptTcpClient()
            $openClients.Add($c)
            $s = $c.GetStream()
            $r = New-Object System.IO.StreamReader($s, [System.Text.Encoding]::ASCII, $false, 1024, $true)
            $reqLines = @()
            while ($null -ne ($l = $r.ReadLine()) -and $l -ne '') {
                $reqLines += $l
            }

            if ($reqLines[0] -match '^HEAD ') {
                $hdr = "HTTP/1.1 200 OK`r`nContent-Type: application/octet-stream`r`nAccept-Ranges: bytes`r`nContent-Length: 1048576`r`nConnection: close`r`n`r`n"
                $hdrBytes = [System.Text.Encoding]::ASCII.GetBytes($hdr)
                $s.Write($hdrBytes, 0, $hdrBytes.Length)
                $s.Flush()
                $c.Close()
            } elseif ($reqLines | Where-Object { $_ -match '^Range:\s*bytes=0-1\b' }) {
                $hdr = "HTTP/1.1 206 Partial Content`r`nContent-Type: application/octet-stream`r`nAccept-Ranges: bytes`r`nContent-Range: bytes 0-1/1048576`r`nContent-Length: 2`r`nConnection: close`r`n`r`n"
                $hdrBytes = [System.Text.Encoding]::ASCII.GetBytes($hdr)
                $s.Write($hdrBytes, 0, $hdrBytes.Length)
                $s.Write([byte[]]@(65, 66), 0, 2)
                $s.Flush()
                $c.Close()
            } else {
                $hdr = "HTTP/1.1 206 Partial Content`r`nContent-Type: application/octet-stream`r`nAccept-Ranges: bytes`r`nContent-Range: bytes 0-1048575/1048576`r`nContent-Length: 1048576`r`nConnection: keep-alive`r`n`r`n"
                $hdrBytes = [System.Text.Encoding]::ASCII.GetBytes($hdr)
                $s.Write($hdrBytes, 0, $hdrBytes.Length)
                $chunk = New-Object byte[] 65536
                $s.Write($chunk, 0, $chunk.Length)
                $s.Flush()
                $transferStarted = $true
            }
        }
        $filesInCancelDir = @(Get-ChildItem -Path $CancelDir -Force -ErrorAction SilentlyContinue)
        if ($transferStarted -and $filesInCancelDir.Count -gt 0) {
            Write-Host "  [Cancel Test] Active download file(s) in $CancelDir : $($filesInCancelDir.Name -join ', ')" -ForegroundColor Gray
            break
        }
        Start-Sleep -Milliseconds 50
    }

    # Simulate user pressing Ctrl+C (Stop job mid-download, which triggers finally {} block)
    Write-Host "  [Cancel Test] Stopping job mid-download (simulating Ctrl+C)..." -ForegroundColor Gray
    Stop-Job -Job $cancelJob
}
finally {
    Stop-Job -Job $cancelJob -ErrorAction SilentlyContinue
    Remove-Job -Job $cancelJob -Force -ErrorAction SilentlyContinue
    foreach ($c in $openClients) { try { $c.Close() } catch {} }
    $slowListener.Stop()
}

# Verify that any .tmp or partial file in $CancelDir is unlocked and can be removed cleanly
Remove-Item -Path $CancelDir -Recurse -Force -ErrorAction Stop

# Verify that the canceled job is no longer in the active Delivery Optimization queue
$RemoveJobScript = Join-Path $ScriptDir "Remove-DODownloadJob.ps1"
$remainingJobs = @(& $RemoveJobScript -DisplayName "DO Slow Cancel Test" -ListOnly)
if ($remainingJobs.Count -gt 0) {
    throw "Test 4 FAILED: Canceled job still exists in Delivery Optimization queue!"
}
Write-Host "[PASS] Mid-download cancellation cleanly aborted DoSvc job (0 remaining in queue) and unlocked/removed temp files!`n" -ForegroundColor Green

Write-Host "=== All Delivery Optimization Tests Passed! ===" -ForegroundColor Cyan
