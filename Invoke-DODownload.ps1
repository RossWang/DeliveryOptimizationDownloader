<#
.SYNOPSIS
    Downloads a file (or byte range) using the Windows Delivery Optimization (DoSvc) COM API.

.PARAMETER Uri
    The remote HTTP/HTTPS URL of the file to download.

.PARAMETER OutFile
    The local destination file path.

.PARAMETER DisplayName
    Optional friendly name for the Delivery Optimization job.

.PARAMETER Background
    Switch to run the download with background priority instead of the default foreground priority.

.PARAMETER RangeOffset
    Optional starting byte offset for a partial range download. Must be paired with -RangeLength.

.PARAMETER RangeLength
    Optional number of bytes to download starting at -RangeOffset.

.PARAMETER Headers
    Optional hashtable or dictionary of custom HTTP request headers (e.g. @{ 'Authorization' = 'Bearer token'; 'X-Custom' = 'Value' }).

.EXAMPLE
    .\Invoke-DODownload.ps1 -Uri "https://example.com/file.zip" -OutFile ".\file.zip"

.EXAMPLE
    .\Invoke-DODownload.ps1 -Uri "https://example.com/file.zip" -OutFile ".\file.zip" -Headers @{ "Authorization" = "Bearer my-token" }

.EXAMPLE
    .\Invoke-DODownload.ps1 -Uri "https://example.com/file.zip" -OutFile ".\part.bin" -RangeOffset 0 -RangeLength 1024
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Uri,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$OutFile,

    [Parameter(Mandatory = $false)]
    [string]$DisplayName = "PowerShell DO Download",

    [Parameter(Mandatory = $false)]
    [switch]$Background,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [long]::MaxValue)]
    [nullable[uint64]]$RangeOffset = $null,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [long]::MaxValue)]
    [nullable[uint64]]$RangeLength = $null,

    [Parameter(Mandatory = $false)]
    [System.Collections.IDictionary]$Headers = $null
)

$ErrorActionPreference = 'Stop'

if (($null -ne $RangeOffset) -xor ($null -ne $RangeLength)) {
    throw "Both -RangeOffset and -RangeLength must be specified together for a partial range download."
}

# Ensure output directory exists and resolve full path
$OutFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)
$parentDir = Split-Path -Path $OutFile -Parent
if ($parentDir -and -not (Test-Path -Path $parentDir)) {
    New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
}

# Compile the C# COM Interop wrapper for IDOManager and IDODownload
if (-not ('DeliveryOptimizationInterop.DODownloader' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace DeliveryOptimizationInterop
{
    public enum DODownloadProperty : int
    {
        Id                       = 0,
        Uri                      = 1,
        ContentId                = 2,
        DisplayName              = 3,
        LocalPath                = 4,
        HttpCustomHeaders        = 5,
        CostPolicy               = 6,
        SecurityFlags            = 7,
        CallbackFreqPercent      = 8,
        CallbackFreqSeconds      = 9,
        NoProgressTimeoutSeconds = 10,
        ForegroundPriority       = 11,
        BlockingMode             = 12
    }

    public enum DODownloadState : int
    {
        Created      = 0,
        Transferring = 1,
        Transferred  = 2,
        Finalized    = 3,
        Aborted      = 4,
        Paused       = 5
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DO_DOWNLOAD_RANGE
    {
        public ulong Offset;
        public ulong Length;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DO_DOWNLOAD_RANGES_INFO
    {
        public uint RangeCount;
        public DO_DOWNLOAD_RANGE FirstRange;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DO_DOWNLOAD_STATUS
    {
        public ulong BytesTotal;
        public ulong BytesTransferred;
        public DODownloadState State;
        public int Error;
        public int ExtendedError;
    }

    [ComImport]
    [Guid("FBBD7FC0-C147-4727-A38D-827EF071EE77")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDODownload
    {
        void Start(IntPtr ranges);
        void Pause();
        void Abort();
        void FinalizeDownload();
        void GetStatus(out DO_DOWNLOAD_STATUS status);
        void GetProperty(DODownloadProperty propId, [MarshalAs(UnmanagedType.Struct)] out object propVal);
        void SetProperty(DODownloadProperty propId, [In, MarshalAs(UnmanagedType.Struct)] ref object propVal);
    }

    [ComImport]
    [Guid("400E2D4A-1431-4C1A-A748-39CA472CFDB1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDOManager
    {
        void CreateDownload([MarshalAs(UnmanagedType.Interface)] out IDODownload download);
        void EnumDownloads(IntPtr category, out IntPtr ppEnum);
    }

    public sealed class DODownloader : IDisposable
    {
        private static readonly Guid CLSID_DeliveryOptimization = new Guid("5b99fa76-721c-423c-adac-56d03c8a8007");

        private IDOManager _manager;
        private IDODownload _download;

        [DllImport("ole32.dll")]
        private static extern int CoSetProxyBlanket(
            IntPtr pProxy,
            uint dwAuthnSvc,
            uint dwAuthzSvc,
            [MarshalAs(UnmanagedType.LPWStr)] string pServerPrincName,
            uint dwAuthnLevel,
            uint dwImpLevel,
            IntPtr pAuthInfo,
            uint dwCapabilities);

        private static void EnableStaticCloaking(object comObject, Type interfaceType)
        {
            const uint RPC_C_AUTHN_DEFAULT = 0xFFFFFFFF;
            const uint RPC_C_AUTHZ_NONE = 0;
            const uint RPC_C_AUTHN_LEVEL_DEFAULT = 0;
            const uint RPC_C_IMP_LEVEL_IMPERSONATE = 3;
            const uint EOAC_STATIC_CLOAKING = 0x20;

            IntPtr pItf = Marshal.GetComInterfaceForObject(comObject, interfaceType);
            try
            {
                int hr = CoSetProxyBlanket(
                    pItf,
                    RPC_C_AUTHN_DEFAULT,
                    RPC_C_AUTHZ_NONE,
                    null,
                    RPC_C_AUTHN_LEVEL_DEFAULT,
                    RPC_C_IMP_LEVEL_IMPERSONATE,
                    IntPtr.Zero,
                    EOAC_STATIC_CLOAKING);

                if (hr < 0)
                    Marshal.ThrowExceptionForHR(hr);
            }
            finally
            {
                if (pItf != IntPtr.Zero)
                    Marshal.Release(pItf);
            }
        }

        public DODownloader()
        {
            Type comType = Type.GetTypeFromCLSID(CLSID_DeliveryOptimization, true);
            object rawMgr = Activator.CreateInstance(comType);
            EnableStaticCloaking(rawMgr, typeof(IDOManager));
            _manager = (IDOManager)rawMgr;

            _manager.CreateDownload(out _download);
            EnableStaticCloaking(_download, typeof(IDODownload));
        }

        public void SetProperty(DODownloadProperty prop, object value)
        {
            _download.SetProperty(prop, ref value);
        }

        public object GetProperty(DODownloadProperty prop)
        {
            object val;
            _download.GetProperty(prop, out val);
            return val;
        }

        public void StartFullDownload()
        {
            DO_DOWNLOAD_RANGES_INFO emptyRanges = new DO_DOWNLOAD_RANGES_INFO();
            emptyRanges.RangeCount = 0;

            IntPtr pRanges = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(DO_DOWNLOAD_RANGES_INFO)));
            try
            {
                Marshal.StructureToPtr(emptyRanges, pRanges, false);
                _download.Start(pRanges);
            }
            finally
            {
                Marshal.FreeHGlobal(pRanges);
            }
        }

        public void StartRangeDownload(ulong offset, ulong length)
        {
            DO_DOWNLOAD_RANGES_INFO rangeInfo = new DO_DOWNLOAD_RANGES_INFO();
            rangeInfo.RangeCount = 1;
            rangeInfo.FirstRange.Offset = offset;
            rangeInfo.FirstRange.Length = length;

            IntPtr pRanges = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(DO_DOWNLOAD_RANGES_INFO)));
            try
            {
                Marshal.StructureToPtr(rangeInfo, pRanges, false);
                _download.Start(pRanges);
            }
            finally
            {
                Marshal.FreeHGlobal(pRanges);
            }
        }

        public void Resume()
        {
            _download.Start(IntPtr.Zero);
        }

        public void Pause()
        {
            _download.Pause();
        }

        public void Abort()
        {
            if (_download != null)
            {
                try { _download.Abort(); } catch { }
            }
        }

        public void FinalizeDownload()
        {
            _download.FinalizeDownload();
        }

        public DO_DOWNLOAD_STATUS GetStatus()
        {
            DO_DOWNLOAD_STATUS status;
            _download.GetStatus(out status);
            return status;
        }

        public void Dispose()
        {
            if (_download != null)
            {
                Marshal.ReleaseComObject(_download);
                _download = null;
            }
            if (_manager != null)
            {
                Marshal.ReleaseComObject(_manager);
                _manager = null;
            }
        }
    }
}
"@
}

# 1. Instantiate DODownloader wrapper
Write-Verbose "Connecting to Delivery Optimization service (DoSvc)..."
$downloader = New-Object DeliveryOptimizationInterop.DODownloader
$downloadCompleted = $false

try {
    # 2. Configure Download Properties
    $downloader.SetProperty('DisplayName', [string]$DisplayName)
    $downloader.SetProperty('Uri', [string]$Uri)
    $downloader.SetProperty('LocalPath', [string]$OutFile)
    $downloader.SetProperty('ForegroundPriority', [bool](-not $Background.IsPresent))

    if ($null -ne $Headers -and $Headers.Count -gt 0) {
        $headerLines = foreach ($key in $Headers.Keys) {
            "{0}: {1}" -f $key, $Headers[$key]
        }
        $rawHeaders = ($headerLines -join "`r`n") + "`r`n"
        $downloader.SetProperty('HttpCustomHeaders', [string]$rawHeaders)
    }

    $jobId = $downloader.GetProperty('Id')

    # 3. Start the Download (Full or Range)
    Write-Host "Starting Delivery Optimization download (Job ID: $jobId)..." -ForegroundColor Cyan
    Write-Host "  Source : $Uri"
    Write-Host "  Target : $OutFile"
    if ($null -ne $Headers -and $Headers.Count -gt 0) {
        Write-Host "  Headers: $($Headers.Keys -join ', ')"
    }

    if ($null -ne $RangeOffset) {
        Write-Host "  Range  : Offset=$RangeOffset, Length=$RangeLength bytes"
        $downloader.StartRangeDownload($RangeOffset, $RangeLength)
    } else {
        $downloader.StartFullDownload()
    }

    # 4. Monitor Progress until Transferred or Error
    do {
        Start-Sleep -Milliseconds 250
        $status = $downloader.GetStatus()

        if ($status.Error -ne 0) {
            $errHex = "0x{0:X8}" -f $status.Error
            $extHex = "0x{0:X8}" -f $status.ExtendedError
            throw "Delivery Optimization error: $errHex (ExtendedError: $extHex)"
        }

        if ($status.BytesTotal -gt 0) {
            $pct = [math]::Min(100, [int](($status.BytesTransferred * 100) / $status.BytesTotal))
            $kbTransferred = [math]::Round($status.BytesTransferred / 1KB, 2)
            $kbTotal       = [math]::Round($status.BytesTotal / 1KB, 2)
            Write-Progress -Activity $DisplayName -Status "$($status.State): $kbTransferred KB / $kbTotal KB ($pct%)" -PercentComplete $pct
        } else {
            Write-Progress -Activity $DisplayName -Status "$($status.State): Connecting..." -PercentComplete 0
        }

    } while ($status.State -in @(
        [DeliveryOptimizationInterop.DODownloadState]::Created,
        [DeliveryOptimizationInterop.DODownloadState]::Transferring,
        [DeliveryOptimizationInterop.DODownloadState]::Paused
    ))

    Write-Progress -Activity $DisplayName -Completed

    # 5. Finalize the download so DoSvc commits the file and removes the job
    if ($status.State -eq [DeliveryOptimizationInterop.DODownloadState]::Transferred) {
        $downloader.FinalizeDownload()
        $downloadCompleted = $true
        Write-Host "Download completed and finalized ($($status.BytesTransferred) bytes): $OutFile" -ForegroundColor Green
    } else {
        throw "Download ended in unexpected state: $($status.State)"
    }
}
finally {
    Write-Progress -Activity $DisplayName -Completed
    if ($null -ne $downloader) {
        if (-not $downloadCompleted) {
            $downloader.Abort()
        }
        $downloader.Dispose()
    }
}
