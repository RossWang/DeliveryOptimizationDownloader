# Delivery Optimization Downloader

A powershell script to download files using [Delivery Optimization](https://learn.microsoft.com/en-us/windows/deployment/do/waas-delivery-optimization#types-of-download-content-supported-by-delivery-optimization).

The benefit is the ability to download from multiple different addresses at the same time. That means different servers, paths, locations or providers.

While other multi-threaded parallel download tools usually download from the same IP address.

## Usage

Download a file

```
.\Invoke-DODownload.ps1 "https://example.com/file.zip" ".\file.zip"
```

Download a file with custom headers

```
.\Invoke-DODownload.ps1 -Uri "https://example.com/file.zip" -OutFile ".\file.zip" -Headers @{ "Authorization" = "Bearer my-token" }
```

Lists all active Delivery Optimization jobs without removing them.

```
.\Remove-DODownloadJob.ps1 -ListOnly -All
```

Removes any Delivery Optimization job targeting the specified path.

```
.\Remove-DODownloadJob.ps1 -LocalPath ".\file.zip"
```

## Speed

- 2x slower on low latency link
- 10x slower on high latency link
- 100x slower when dealing with rare server implementation
- Sometimes it's faster due to the spreading of servers and paths

[Share your experiences!](https://github.com/RossWang/DeliveryOptimizationDownloader/discussions/1)
