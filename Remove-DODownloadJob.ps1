<#
.SYNOPSIS
    Lists or removes (aborts) active/orphaned Windows Delivery Optimization download jobs.

.PARAMETER DisplayName
    Filter jobs by DisplayName (substring or wildcard match). Defaults to "PowerShell DO Download".

.PARAMETER JobId
    Filter by a specific Delivery Optimization Job ID (GUID).

.PARAMETER LocalPath
    Filter by the target local file path.

.PARAMETER All
    Switch to remove all enumerated Delivery Optimization jobs accessible to the current user.

.PARAMETER ListOnly
    Switch to list matching jobs without aborting/removing them.

.EXAMPLE
    .\Remove-DODownloadJob.ps1
    Removes all jobs matching DisplayName "PowerShell DO Download".

.EXAMPLE
    .\Remove-DODownloadJob.ps1 -ListOnly -All
    Lists all active Delivery Optimization jobs without removing them.

.EXAMPLE
    .\Remove-DODownloadJob.ps1 -LocalPath ".\file.zip"
    Removes any Delivery Optimization job targeting the specified path.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$DisplayName = "PowerShell DO Download",

    [Parameter(Mandatory = $false)]
    [string]$JobId,

    [Parameter(Mandatory = $false)]
    [string]$LocalPath,

    [Parameter(Mandatory = $false)]
    [switch]$All,

    [Parameter(Mandatory = $false)]
    [switch]$ListOnly
)

$ErrorActionPreference = 'Stop'

if (-not ('DeliveryOptimizationCleanup.DOJobManager' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace DeliveryOptimizationCleanup
{
    public enum DODownloadProperty : int
    {
        Id          = 0,
        Uri         = 1,
        ContentId   = 2,
        DisplayName = 3,
        LocalPath   = 4
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
    public struct DO_DOWNLOAD_STATUS
    {
        public ulong BytesTotal;
        public ulong BytesTransferred;
        public DODownloadState State;
        public int Error;
        public int ExtendedError;
    }

    public sealed class DOJobInfo
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Uri { get; set; }
        public string LocalPath { get; set; }
        public DODownloadState State { get; set; }
        public ulong BytesTransferred { get; set; }
        public ulong BytesTotal { get; set; }
        public bool Aborted { get; set; }
    }

    [ComImport]
    [Guid("00000100-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumUnknown
    {
        [PreserveSig]
        int Next(uint celt, [Out, MarshalAs(UnmanagedType.IUnknown)] out object rgelt, out uint pceltFetched);
        [PreserveSig]
        int Skip(uint celt);
        [PreserveSig]
        int Reset();
        void Clone([MarshalAs(UnmanagedType.Interface)] out IEnumUnknown ppenum);
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
        [PreserveSig]
        int EnumDownloads(IntPtr category, [MarshalAs(UnmanagedType.Interface)] out IEnumUnknown ppEnum);
    }

    public static class DOJobManager
    {
        private static readonly Guid CLSID_DeliveryOptimization = new Guid("5b99fa76-721c-423c-adac-56d03c8a8007");

        [DllImport("ole32.dll")]
        private static extern int CoSetProxyBlanket(
            IntPtr pProxy, uint dwAuthnSvc, uint dwAuthzSvc,
            [MarshalAs(UnmanagedType.LPWStr)] string pServerPrincName,
            uint dwAuthnLevel, uint dwImpLevel, IntPtr pAuthInfo, uint dwCapabilities);

        private static void EnableCloaking(object obj, Type itfType)
        {
            IntPtr p = Marshal.GetComInterfaceForObject(obj, itfType);
            try
            {
                int hr = CoSetProxyBlanket(p, 0xFFFFFFFF, 0, null, 0, 3, IntPtr.Zero, 0x20);
                if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            }
            finally
            {
                if (p != IntPtr.Zero) Marshal.Release(p);
            }
        }

        private static string TryGetStringProp(IDODownload job, DODownloadProperty prop)
        {
            try
            {
                object val;
                job.GetProperty(prop, out val);
                return val as string ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static List<DOJobInfo> ProcessJobs(
            string filterDisplayName,
            string filterJobId,
            string filterLocalPath,
            bool matchAll,
            bool listOnly)
        {
            var results = new List<DOJobInfo>();
            Type comType = Type.GetTypeFromCLSID(CLSID_DeliveryOptimization, true);
            object rawMgr = Activator.CreateInstance(comType);
            EnableCloaking(rawMgr, typeof(IDOManager));
            IDOManager mgr = (IDOManager)rawMgr;

            try
            {
                IEnumUnknown enumUnk;
                int hr = mgr.EnumDownloads(IntPtr.Zero, out enumUnk);
                // 0x80D02005 means no download jobs exist in the queue
                if (hr < 0 || enumUnk == null) return results;
                EnableCloaking(enumUnk, typeof(IEnumUnknown));

                try
                {
                    object item;
                    uint fetched;
                    while (enumUnk.Next(1, out item, out fetched) == 0 && fetched == 1 && item != null)
                    {
                        IDODownload job = item as IDODownload;
                        if (job == null) continue;

                        try
                        {
                            EnableCloaking(job, typeof(IDODownload));
                            var info = new DOJobInfo
                            {
                                Id          = TryGetStringProp(job, DODownloadProperty.Id),
                                DisplayName = TryGetStringProp(job, DODownloadProperty.DisplayName),
                                Uri         = TryGetStringProp(job, DODownloadProperty.Uri),
                                LocalPath   = TryGetStringProp(job, DODownloadProperty.LocalPath)
                            };

                            try
                            {
                                DO_DOWNLOAD_STATUS st;
                                job.GetStatus(out st);
                                info.State            = st.State;
                                info.BytesTransferred = st.BytesTransferred;
                                info.BytesTotal       = st.BytesTotal;
                            }
                            catch { }

                            bool isMatch = matchAll;
                            if (!isMatch)
                            {
                                if (!string.IsNullOrEmpty(filterJobId))
                                    isMatch = string.Equals(info.Id, filterJobId, StringComparison.OrdinalIgnoreCase);
                                else if (!string.IsNullOrEmpty(filterLocalPath))
                                    isMatch = string.Equals(info.LocalPath, filterLocalPath, StringComparison.OrdinalIgnoreCase);
                                else if (!string.IsNullOrEmpty(filterDisplayName))
                                    isMatch = info.DisplayName.IndexOf(filterDisplayName, StringComparison.OrdinalIgnoreCase) >= 0;
                            }

                            if (isMatch)
                            {
                                if (!listOnly)
                                {
                                    job.Abort();
                                    info.Aborted = true;
                                    info.State = DODownloadState.Aborted;
                                }
                                results.Add(info);
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(job);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(enumUnk);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(mgr);
            }

            return results;
        }
    }
}
"@
}

$resolvedPath = if (-not [string]::IsNullOrWhiteSpace($LocalPath)) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LocalPath)
} else {
    ""
}

$jobs = [DeliveryOptimizationCleanup.DOJobManager]::ProcessJobs(
    $DisplayName,
    $JobId,
    $resolvedPath,
    $All.IsPresent,
    $ListOnly.IsPresent
)

if ($jobs.Count -eq 0) {
    Write-Host "No matching Delivery Optimization jobs found." -ForegroundColor Yellow
} else {
    foreach ($j in $jobs) {
        if ($ListOnly) {
            Write-Host ("[Job] ID={0} | State={1} | Name='{2}' | Path='{3}'" -f $j.Id, $j.State, $j.DisplayName, $j.LocalPath) -ForegroundColor Cyan
        } else {
            Write-Host ("[Removed] ID={0} | Name='{1}' | Path='{2}'" -f $j.Id, $j.DisplayName, $j.LocalPath) -ForegroundColor Green
        }
    }
    $jobs
}
