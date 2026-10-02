# Installs the VB-Audio Virtual Cable driver on a headless Windows runner, for the
# device backed tests in tests/device/.
#
# Ported from CamillaDSP, testscripts/e2e/wasapi/install_vbcable.ps1 in
# https://github.com/HEnquist/camilladsp, with the download moved into the script
# and pinned to a known archive.
#
# Does what `devcon install <inf> <hwid>` does, through SetupAPI directly, so this
# needs neither devcon.exe nor a third party action that bundles it, and no reboot.
# The driver's signer is added to TrustedPublisher first, since the install is
# otherwise stopped by a prompt asking whether to trust the publisher, which nobody
# is there to answer.
#
# VB-CABLE is donationware from VB-Audio: free for personal and evaluation use,
# with a licence required for professional or commercial use, and no redistribution
# without an agreement. See https://vb-audio.com/Cable/. The archive is therefore
# downloaded at run time and deliberately not committed to this repository.
#
# Usage: install_vbcable.ps1 [-Instances <n>]

param(
    [int]$Instances = 1
)

$ErrorActionPreference = 'Stop'

# Pinned, and the hash is checked before anything is unpacked. This installs a
# kernel driver, so a silently changed archive must fail rather than be trusted.
# Driver pack 45, DriverVer 10/07/2024 3.3.1.7.
$Url = 'https://download.vb-audio.com/Download_CABLE/VBCABLE_Driver_Pack45.zip'
$Sha256 = 'B950E39F01AF1D04EA623C8F6D8EB9B6EA5C477C637295FABF20631C85116BFB'
$InfName = 'vbMmeCable64_win10.inf'
$HardwareId = 'VBAudioVACWDM'

$work = if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } else { $env:TEMP }
$zip = Join-Path $work 'vbcable.zip'
$dir = Join-Path $work 'vbcable'

# vb-audio.com is a single small host, so a transient failure is worth a retry.
for ($attempt = 1; $attempt -le 3; $attempt++) {
    try {
        Invoke-WebRequest $Url -OutFile $zip
        break
    } catch {
        Write-Host "download attempt $attempt failed: $($_.Exception.Message)"
        if ($attempt -eq 3) { throw }
        Start-Sleep -Seconds (5 * $attempt)
    }
}

$hash = (Get-FileHash $zip -Algorithm SHA256).Hash
if ($hash -ne $Sha256) {
    throw "the archive hash is $hash, expected $Sha256. Refusing to install it."
}
Write-Host "downloaded $Url, hash verified"

Expand-Archive $zip -DestinationPath $dir -Force
$Inf = Join-Path (Resolve-Path $dir) $InfName
if (-not (Test-Path $Inf)) { throw "$InfName is not in the archive" }

# Trust every certificate the catalogs are signed with, rather than a .cer kept in
# the repo, so a new driver pack brings its own.
foreach ($cat in Get-ChildItem $dir -Filter *.cat) {
    $sig = Get-AuthenticodeSignature $cat.FullName
    Write-Host "$($cat.Name): $($sig.Status), $($sig.SignerCertificate.Subject)"
    if ($sig.SignerCertificate) {
        $cer = Join-Path $work "$($cat.BaseName).cer"
        Export-Certificate -Cert $sig.SignerCertificate -FilePath $cer | Out-Null
        certutil -f -addstore TrustedPublisher $cer | Out-Null
    }
}

Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

public static class DevInstall {
    const int DICD_GENERATE_ID = 1;
    const int SPDRP_HARDWAREID = 1;
    const int DIF_REGISTERDEVICE = 0x19;
    const int INSTALLFLAG_FORCE = 1;
    static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

    [StructLayout(LayoutKind.Sequential)]
    struct SP_DEVINFO_DATA {
        public int cbSize;
        public Guid ClassGuid;
        public int DevInst;
        public IntPtr Reserved;
    }

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool SetupDiGetINFClass(string infName, out Guid classGuid,
        StringBuilder className, int classNameSize, out int requiredSize);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern IntPtr SetupDiCreateDeviceInfoList(ref Guid classGuid, IntPtr hwndParent);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool SetupDiCreateDeviceInfo(IntPtr set, string deviceName, ref Guid classGuid,
        string description, IntPtr hwndParent, int flags, ref SP_DEVINFO_DATA data);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool SetupDiSetDeviceRegistryProperty(IntPtr set, ref SP_DEVINFO_DATA data,
        int property, byte[] buffer, int size);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiCallClassInstaller(int function, IntPtr set, ref SP_DEVINFO_DATA data);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiDestroyDeviceInfoList(IntPtr set);

    [DllImport("newdev.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool UpdateDriverForPlugAndPlayDevices(IntPtr hwndParent, string hardwareId,
        string infPath, int flags, out bool rebootRequired);

    static void Check(bool ok, string what) {
        if (!ok) throw new Win32Exception(Marshal.GetLastWin32Error(), what);
    }

    // Creates a root enumerated device node with the hardware id, then points the
    // driver at every node with that id. Returns whether Windows asks for a reboot.
    public static bool Install(string inf, string hardwareId) {
        Guid classGuid;
        int required;
        StringBuilder className = new StringBuilder(32);
        Check(SetupDiGetINFClass(inf, out classGuid, className, className.Capacity, out required),
            "SetupDiGetINFClass");

        IntPtr set = SetupDiCreateDeviceInfoList(ref classGuid, IntPtr.Zero);
        if (set == INVALID_HANDLE_VALUE) Check(false, "SetupDiCreateDeviceInfoList");
        try {
            SP_DEVINFO_DATA data = new SP_DEVINFO_DATA();
            data.cbSize = Marshal.SizeOf(data);
            Check(SetupDiCreateDeviceInfo(set, className.ToString(), ref classGuid, null,
                IntPtr.Zero, DICD_GENERATE_ID, ref data), "SetupDiCreateDeviceInfo");
            // REG_MULTI_SZ, so the list ends with an extra null.
            byte[] id = Encoding.Unicode.GetBytes(hardwareId + "\0\0");
            Check(SetupDiSetDeviceRegistryProperty(set, ref data, SPDRP_HARDWAREID, id, id.Length),
                "SetupDiSetDeviceRegistryProperty");
            Check(SetupDiCallClassInstaller(DIF_REGISTERDEVICE, set, ref data),
                "SetupDiCallClassInstaller");
        } finally {
            SetupDiDestroyDeviceInfoList(set);
        }

        bool reboot;
        Check(UpdateDriverForPlugAndPlayDevices(IntPtr.Zero, hardwareId, inf, INSTALLFLAG_FORCE,
            out reboot), "UpdateDriverForPlugAndPlayDevices");
        return reboot;
    }
}
'@

function Get-CableEndpoints {
    Get-PnpDevice -Class AudioEndpoint -Status OK -ErrorAction SilentlyContinue |
        Where-Object FriendlyName -like '*(VB-Audio Virtual Cable)'
}

$before = @(Get-CableEndpoints).Count
for ($i = 1; $i -le $Instances; $i++) {
    $reboot = [DevInstall]::Install($Inf, $HardwareId)
    Write-Host "instance $i installed, reboot requested: $reboot"
}

# The install returns before the audio endpoints exist, and a client that looks for
# them too early finds nothing. Each instance brings three: two playback, one capture.
$start = Get-Date
$wanted = $before + 3 * $Instances
while (@(Get-CableEndpoints).Count -lt $wanted) {
    if (((Get-Date) - $start).TotalSeconds -gt 60) {
        Get-CableEndpoints | Format-Table Status, FriendlyName -AutoSize
        Write-Host "::error::only $(@(Get-CableEndpoints).Count) of $wanted endpoints after 60 s"
        exit 1
    }
    Start-Sleep -Milliseconds 200
}
Write-Host "$wanted endpoints after $([int]((Get-Date) - $start).TotalMilliseconds) ms"
Get-CableEndpoints | Format-Table Status, FriendlyName -AutoSize

# Record which driver version ended up installed, so the log says so on a failure.
pnputil /enum-drivers | Select-String -Pattern 'vbMmeCable', 'VB-Audio' -Context 2, 4
