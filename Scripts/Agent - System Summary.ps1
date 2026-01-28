Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Custom Panel class with rounded corners and border
Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public class RoundedPanel : Panel
{
    private Color borderColor = Color.FromArgb(220, 220, 220);
    private int borderRadius = 6;
    private int borderWidth = 1;

    public Color BorderColor
    {
        get { return borderColor; }
        set { borderColor = value; Invalidate(); }
    }

    public int BorderRadius
    {
        get { return borderRadius; }
        set { borderRadius = value; Invalidate(); }
    }

    public int BorderWidth
    {
        get { return borderWidth; }
        set { borderWidth = value; Invalidate(); }
    }

    public RoundedPanel()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        
        Rectangle rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
        
        using (GraphicsPath path = GetRoundedRectangle(rect, borderRadius))
        {
            // Set clipping region for rounded corners
            Region = new Region(path);
            
            // Draw border
            using (Pen pen = new Pen(borderColor, borderWidth))
            {
                g.DrawPath(pen, path);
            }
        }
    }

    private GraphicsPath GetRoundedRectangle(Rectangle rect, int radius)
    {
        GraphicsPath path = new GraphicsPath();
        
        int diameter = radius * 2;
        
        path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
        path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
        path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        
        return path;
    }
}
"@

# Compatibility function for Windows 7 (PowerShell 2.0) support
# Tries Get-CimInstance first (PowerShell 3.0+), falls back to Get-WmiObject (PowerShell 2.0+)
function Get-WmiCompat {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ClassName,
        
        [Parameter(Mandatory=$false)]
        [string]$Filter
    )
    
    # Try Get-CimInstance first (PowerShell 3.0+)
    if (Get-Command Get-CimInstance -ErrorAction SilentlyContinue) {
        if ($Filter) {
            return Get-CimInstance -ClassName $ClassName -Filter $Filter -ErrorAction SilentlyContinue
        } else {
            return Get-CimInstance -ClassName $ClassName -ErrorAction SilentlyContinue
        }
    }
    # Fallback to Get-WmiObject (PowerShell 2.0+)
    else {
        if ($Filter) {
            return Get-WmiObject -Class $ClassName -Filter $Filter -ErrorAction SilentlyContinue
        } else {
            return Get-WmiObject -Class $ClassName -ErrorAction SilentlyContinue
        }
    }
}

function Get-SystemInfo {
    # Device Info
    $computerName = [string]$env:COMPUTERNAME
    $userName     = [string]$env:USERNAME
    
    try {
        $compSys = Get-WmiCompat -ClassName Win32_ComputerSystem
        $domain = if ($compSys) { [string]$compSys.Domain } else { "Unavailable" }
    } catch {
        $domain = "Unavailable"
    }

    # OS Info
    try {
        $os = Get-WmiCompat -ClassName Win32_OperatingSystem
        if (-not $os) { throw "Failed to get OS info" }
        $osVersion = [string]$os.Caption
        $osArch = if ($os.OSArchitecture) { $os.OSArchitecture } else { "Unavailable" }
        # Handle date conversion for WMI compatibility (Windows 7/PowerShell 2.0)
        if ($os.LastBootUpTime -is [DateTime]) {
            $lastReboot = $os.LastBootUpTime
        } else {
            # Get-WmiObject returns ManagementDateTimeConverter string, need to convert
            $lastReboot = [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
        }
        $uptimeSpan = (Get-Date) - $lastReboot
        $uptime = "{0}D {1}H {2}M" -f $uptimeSpan.Days, $uptimeSpan.Hours, $uptimeSpan.Minutes
        
        # Detailed OS version info
        $osBuildNumber = $os.BuildNumber
        $osVersionNumber = $os.Version
        $osSerialNumber = $os.SerialNumber
        $osDetailedInfo = "OS Name: $osVersion`nArchitecture: $osArch`nVersion: $osVersionNumber`nBuild Number: $osBuildNumber`nSerial Number: $osSerialNumber"
    } catch {
        $osVersion = "Unavailable"
        $osArch = "Unavailable"
        $lastReboot = Get-Date
        $uptime = "Unavailable"
        $osDetailedInfo = "OS information unavailable"
    }

    # Hardware Info
    try {
        $totalRAMGB = if ($compSys) { 
            [math]::Round($compSys.TotalPhysicalMemory / 1GB, 2).ToString() + " GB" 
        } else { 
            "Unavailable" 
        }
    } catch {
        $totalRAMGB = "Unavailable"
    }

    try {
        $cpuObj = Get-WmiCompat -ClassName Win32_Processor | Select-Object -First 1
        $cpu = if ($cpuObj) { $cpuObj.Name.Trim() } else { "Unavailable" }
        
        # Get total cores (physical cores across all processors)
        $allProcessors = Get-WmiCompat -ClassName Win32_Processor
        $totalCores = ($allProcessors | Measure-Object -Property NumberOfCores -Sum).Sum
        $totalLogicalProcessors = ($allProcessors | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
        $coresDisplay = if ($totalCores) { "$totalCores Cores ($totalLogicalProcessors Logical)" } else { "Unavailable" }
    } catch {
        $cpu = "Unavailable"
        $coresDisplay = "Unavailable"
    }

    try {
        $gpuObjects = Get-WmiCompat -ClassName Win32_VideoController | Where-Object { $_.Name -and $_.Name -notmatch 'Microsoft|Remote|Virtual|Basic Display' }
        $gpus = @()
        if ($gpuObjects) {
            foreach ($gpuObj in $gpuObjects) {
                $gpuName = $gpuObj.Name.Trim()
                if ($gpuName -and $gpuName -ne "") {
                    $gpus += $gpuName
                }
            }
        }
        if ($gpus.Count -eq 0) {
            $gpus = @("Unavailable")
        }
    } catch {
        $gpus = @("Unavailable")
    }

    try {
        $manufacturer = if ($compSys.Manufacturer) { $compSys.Manufacturer.Trim() } else { "" }
        $model = if ($compSys.Model -and $compSys.Model -notmatch "System Product Name") { $compSys.Model.Trim() } else { "" }
        $sysModel = "$manufacturer $model".Trim()
        if (-not $sysModel) { $sysModel = "Unavailable" }
    } catch {
        $sysModel = "Unavailable"
    }

    # Disk Info
    $disks = @()
    try {
        $diskObjects = Get-WmiCompat -ClassName Win32_LogicalDisk -Filter "DriveType=3"
        foreach ($disk in $diskObjects) {
            $label = $disk.DeviceID
            $free = [math]::Round($disk.FreeSpace / 1GB, 2)
            $total = [math]::Round($disk.Size / 1GB, 2)
            $used = $total - $free
            $percentUsed = if ($total -gt 0) { [math]::Round(($used / $total) * 100, 1) } else { 0 }
            $disks += @{
                Label = $label
                Free = $free
                Total = $total
                Used = $used
                PercentUsed = $percentUsed
            }
        }
    } catch {
        $disks = @()
    }

    # Network Info - Get all adapters
    $networkAdapters = @()
    try {
        $allAdapters = Get-WmiCompat -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }
        $physicalAdapters = Get-WmiCompat -ClassName Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true }
        $physicalAdapterIndexes = $physicalAdapters.Index
        
        foreach ($adapter in $allAdapters) {
            $adapterIndex = $adapter.Index
            $adapterName = [string]$adapter.Description
            $ipAddresses = ($adapter.IPAddress | Where-Object { $_ -match '^\d+\.' })
            $ipAddress = if ($ipAddresses) { [string]($ipAddresses | Select-Object -First 1) } else { "" }
            $subnet = [string](($adapter.IPSubnet | Where-Object { $_ -match '^\d+\.' }) | Select-Object -First 1)
            $gateway = [string](($adapter.DefaultIPGateway | Where-Object { $_ -match '^\d+\.' }) | Select-Object -First 1)
            $dnsServers = (($adapter.DNSServerSearchOrder | Where-Object { $_ -match '^\d+\.' }) -join ", ")
            if (-not $dnsServers) { $dnsServers = "Unavailable" }
            
            # Determine if this is a physical adapter
            $isPhysical = $physicalAdapterIndexes -contains $adapterIndex
            
            # Determine if this is likely an active adapter
            $isActive = $false
            if ($gateway) {
                if ($isPhysical) {
                    $isActive = $true
                } elseif (-not ($adapterName -match "Virtual|Tunnel|VPN|Loopback|Teredo|6to4|ISATAP|ZeroTier|WireGuard|Tailscale")) {
                    $isActive = $true
                }
            } elseif ($isPhysical -and -not ($adapterName -match "Virtual|Tunnel|VPN|Loopback|Teredo|6to4|ISATAP|ZeroTier|WireGuard|Tailscale")) {
                $isActive = $true
            }
            
            if ($ipAddress) {
                $networkAdapters += @{
                    Name = $adapterName
                    IPAddress = $ipAddress
                    Subnet = $subnet
                    Gateway = $gateway
                    DNSServers = $dnsServers
                    IsPhysical = $isPhysical
                    IsActive = $isActive
                }
            }
        }
        
        # If no adapter marked as active, mark at least one as active
        $hasActive = ($networkAdapters | Where-Object { $_.IsActive }).Count -gt 0
        if (-not $hasActive -and $networkAdapters.Count -gt 0) {
            $bestAdapter = $networkAdapters | Where-Object { $_.Gateway -and $_.IsPhysical } | Select-Object -First 1
            if (-not $bestAdapter) {
                $bestAdapter = $networkAdapters | Where-Object { $_.IsPhysical } | Select-Object -First 1
            }
            if (-not $bestAdapter) {
                $bestAdapter = $networkAdapters | Where-Object { $_.Gateway } | Select-Object -First 1
            }
            if (-not $bestAdapter) {
                $bestAdapter = $networkAdapters[0]
            }
            if ($bestAdapter) {
                $bestAdapter.IsActive = $true
            }
        }
    } catch {
        $networkAdapters = @()
    }

    return @{
        ComputerName = $computerName
        UserName = $userName
        Domain = $domain
        OSVersion = $osVersion
        OSArch = $osArch
        OSDetailedInfo = $osDetailedInfo
        LastReboot = $lastReboot
        Uptime = $uptime
        TotalRAM = $totalRAMGB
        CPU = $cpu
        Cores = $coresDisplay
        GPUs = $gpus
        SystemModel = $sysModel
        Disks = $disks
        NetworkAdapters = $networkAdapters
    }
}

# Colors
$headerColor = [System.Drawing.Color]::FromArgb(0, 70, 140)      # Blue
$accentColor = [System.Drawing.Color]::FromArgb(255, 140, 0)     # Orange
$labelColor = [System.Drawing.Color]::FromArgb(100, 100, 100)    # Medium gray (consistent, readable)
$valueColor = [System.Drawing.Color]::FromArgb(15, 15, 15)       # Darker near-black for values
$adapterHeaderColor = [System.Drawing.Color]::FromArgb(70, 70, 70)  # Slightly darker for adapter names
$dividerColor = [System.Drawing.Color]::FromArgb(220, 220, 220)  # Light gray for dividers
$backgroundColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$contentColor = [System.Drawing.Color]::White

# Fonts
$headerFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$valueFont = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$labelFont = New-Object System.Drawing.Font("Segoe UI", 7.5, [System.Drawing.FontStyle]::Regular)

# Get system information
$sysInfo = Get-SystemInfo

# Get all active network adapters with valid gateways and DNS servers
$activeAdapters = $sysInfo.NetworkAdapters | Where-Object { 
    $_.IsActive -and 
    $_.Gateway -and 
    $_.Gateway -ne "Unavailable" -and
    $_.Gateway.Trim() -ne "" -and
    $_.Gateway -match '^\d+\.\d+\.\d+\.\d+$' -and
    $_.Gateway -ne "0.0.0.0" -and
    $_.DNSServers -and 
    $_.DNSServers -ne "Unavailable" -and
    $_.DNSServers.Trim() -ne ""
} | Sort-Object { 
    if ($_.IsPhysical) { 1 } else { 2 }
}

# Detect connection type from active adapters
function Get-ConnectionType {
    param([array]$Adapters)
    
    if ($Adapters.Count -eq 0) {
        return "Unknown"
    }
    
    $hasWiFi = $false
    $hasEthernet = $false
    
    foreach ($adapter in $Adapters) {
        $adapterName = $adapter.Name.ToLower()
        
        if ($adapterName -match 'wi-fi|wireless|802\.11') {
            $hasWiFi = $true
        }
        if ($adapterName -match 'ethernet|gigabit') {
            $hasEthernet = $true
        }
    }
    
    if ($hasWiFi -and $hasEthernet) {
        return "Multiple (Wi-Fi + Ethernet)"
    } elseif ($hasWiFi) {
        return "Wi-Fi"
    } elseif ($hasEthernet) {
        return "Ethernet"
    } else {
        return "Unknown"
    }
}

$connectionType = Get-ConnectionType -Adapters $activeAdapters

# Function to set HTML to clipboard (Windows HTML clipboard format)
function Set-ClipboardHtml {
    param([string]$html)
    
    # Create HTML clipboard format with required header (Windows clipboard HTML format)
    $htmlFragment = "<html><body>`r`n<!--StartFragment-->$html<!--EndFragment-->`r`n</body></html>"
    
    $startHtml = 0
    $startFragment = $htmlFragment.IndexOf('<!--StartFragment-->') + 20
    $endFragment = $htmlFragment.IndexOf('<!--EndFragment-->')
    $endHtml = $htmlFragment.Length
    
    $htmlHeader = "Version:0.9`r`n"
    $htmlHeader += "StartHTML:$($startHtml.ToString('0000000000'))`r`n"
    $htmlHeader += "EndHTML:$($endHtml.ToString('0000000000'))`r`n"
    $htmlHeader += "StartFragment:$($startFragment.ToString('0000000000'))`r`n"
    $htmlHeader += "EndFragment:$($endFragment.ToString('0000000000'))`r`n"
    $htmlHeader += "`r`n"
    
    $fullHtml = $htmlHeader + $htmlFragment
    
    # Use DataObject to support both HTML and plain text formats
    $dataObject = New-Object System.Windows.Forms.DataObject
    $dataObject.SetData([System.Windows.Forms.DataFormats]::Html, $fullHtml)
    
    # Also include plain text fallback
    $plainText = $html -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"'
    $dataObject.SetText($plainText)
    
    [System.Windows.Forms.Clipboard]::SetDataObject($dataObject, $true)
}

# Function to format system information for email (HTML)
function Format-SystemInfoForEmail {
    param($SysInfo, $ConnectionType, $ActiveAdapters)
    
    # Escape HTML special characters
    function Escape-Html {
        param([string]$text)
        return $text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
    }
    
    # Format Last Reboot with timezone
    $timezone = [TimeZoneInfo]::Local
    $timezoneAbbr = if ($timezone.Id -match 'Eastern') { if ($timezone.IsDaylightSavingTime($SysInfo.LastReboot)) { 'EDT' } else { 'EST' } }
                    elseif ($timezone.Id -match 'Central') { if ($timezone.IsDaylightSavingTime($SysInfo.LastReboot)) { 'CDT' } else { 'CST' } }
                    elseif ($timezone.Id -match 'Mountain') { if ($timezone.IsDaylightSavingTime($SysInfo.LastReboot)) { 'MDT' } else { 'MST' } }
                    elseif ($timezone.Id -match 'Pacific') { if ($timezone.IsDaylightSavingTime($SysInfo.LastReboot)) { 'PDT' } else { 'PST' } }
                    elseif ($timezone.Id -match 'Alaska') { if ($timezone.IsDaylightSavingTime($SysInfo.LastReboot)) { 'AKDT' } else { 'AKST' } }
                    elseif ($timezone.Id -match 'Hawaii') { 'HST' }
                    else { $timezone.Id }
    $lastRebootFormatted = $SysInfo.LastReboot.ToString("M/d/yy h:mm tt") + " " + $timezoneAbbr
    
    $osDisplay = if ($SysInfo.OSArch -ne "Unavailable" -and $SysInfo.OSVersion -ne "Unavailable") {
        "$(Escape-Html $SysInfo.OSVersion) ($(Escape-Html $SysInfo.OSArch))"
    } elseif ($SysInfo.OSVersion -ne "Unavailable") {
        Escape-Html $SysInfo.OSVersion
    } else {
        "Unavailable"
    }
    
    $generatedDate = Get-Date -Format 'M/d/yyyy h:mm tt'
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 7pt;
            color: #333333;
            background-color: #ffffff;
            margin: 0;
            padding: 4px;
            text-align: left;
        }
        @media (prefers-color-scheme: dark) {
            .main-table th {
                background-color: #00468c !important;
                color: #ffffff !important;
                -webkit-text-fill-color: #ffffff !important;
            }
        }
        .wrapper {
            width: 700px;
            margin: 0;
            text-align: left;
        }
        .main-table {
            width: 700px;
            border-collapse: collapse;
            border: 1px solid #dcdcdc;
            font-size: 7pt;
            margin: 0;
        }
        .main-table th {
            background-color: #00468c !important;
            color: #ffffff !important;
            padding: 4px 6px;
            font-weight: bold;
            text-align: left;
            font-size: 9pt;
        }
        .main-table th * {
            color: #ffffff !important;
        }
        .main-table td {
            padding: 2px 6px;
            border-bottom: 1px solid #e8e8e8;
            vertical-align: top;
            font-size: 7pt;
        }
        .section-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 7pt;
        }
        .section-table td {
            padding: 2px 4px;
            border-bottom: 1px solid #e8e8e8;
            vertical-align: top;
            font-size: 7pt;
        }
        .section-table td:first-child {
            font-weight: bold;
            color: #666666;
            width: 90px;
            background-color: #f8f8f8;
            white-space: nowrap;
        }
        .section-table td:last-child {
            color: #1a1a1a;
            word-wrap: break-word;
            max-width: 200px;
        }
        .section-header {
            background-color: #e0e8f0 !important;
            color: #00468c !important;
            font-weight: bold !important;
            font-size: 8pt !important;
            padding: 5px 8px !important;
            border-top: 3px solid #00468c !important;
            border-bottom: 2px solid #00468c !important;
        }
        .section-header td {
            border-bottom: 2px solid #00468c !important;
            padding: 5px 8px !important;
            letter-spacing: 0.5px;
        }
        .disk-table {
            width: 100%;
            border-collapse: collapse;
            margin: 1px 0;
            font-size: 7pt;
            table-layout: fixed;
        }
        .disk-table th {
            background-color: #f0f0f0;
            color: #666666;
            font-weight: bold;
            padding: 3px 5px;
            text-align: left;
            border: 1px solid #dcdcdc;
            font-size: 7pt;
        }
        .disk-table th:nth-child(1) {
            width: 60px;
        }
        .disk-table th:nth-child(2),
        .disk-table th:nth-child(3),
        .disk-table th:nth-child(4) {
            text-align: right;
            width: 100px;
        }
        .disk-table td {
            padding: 2px 5px;
            border: 1px solid #e8e8e8;
            font-size: 7pt;
        }
        .disk-table td:first-child {
            font-weight: bold;
            width: 60px;
        }
        .disk-table td:nth-child(2),
        .disk-table td:nth-child(3),
        .disk-table td:nth-child(4) {
            text-align: right;
            font-variant-numeric: tabular-nums;
            width: 100px;
        }
        .network-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 7pt;
            table-layout: fixed;
        }
        .network-table th {
            background-color: #f0f0f0;
            color: #666666;
            font-weight: bold;
            padding: 3px 6px;
            text-align: left;
            border: 1px solid #dcdcdc;
            font-size: 7pt;
        }
        .network-table th:nth-child(1) {
            width: 200px;
        }
        .network-table th:nth-child(2) {
            width: 120px;
        }
        .network-table th:nth-child(3) {
            width: 120px;
        }
        .network-table th:nth-child(4) {
            width: 240px;
        }
        .network-table td {
            padding: 3px 6px;
            border: 1px solid #e8e8e8;
            font-size: 7pt;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        .network-table td:nth-child(1) {
            width: 200px;
        }
        .network-table td:nth-child(2) {
            width: 120px;
        }
        .network-table td:nth-child(3) {
            width: 120px;
        }
        .network-table td:nth-child(4) {
            width: 240px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .ip-address {
            color: #ff8c00;
            font-weight: bold;
        }
        .footer {
            padding: 3px 6px;
            text-align: left;
            color: #666666;
            font-size: 6pt;
            margin-top: 2px;
        }
    </style>
</head>
<body style="margin: 0; padding: 4px; text-align: left;">
    <div class="wrapper" style="width: 700px; margin: 0; text-align: left;">
    <table class="main-table" style="width: 700px; margin: 0; border-collapse: collapse;">
        <tr>
            <td colspan="4" style="background-color: #00468c; padding: 4px 6px; font-weight: bold; font-size: 9pt;">
                <span style="color: #ffffff; background-color: #00468c;">SYSTEM SUMMARY ($(Escape-Html $generatedDate))</span>
            </td>
        </tr>
        <tr>
            <td colspan="2" style="width: 50%; padding: 4px; vertical-align: top;">
                <table class="section-table" style="width: 100%;">
                    <tr class="section-header">
                        <td colspan="2">DEVICE INFORMATION</td>
                    </tr>
                    <tr><td>Computer:</td><td>$(Escape-Html $SysInfo.ComputerName)</td></tr>
                    <tr><td>User:</td><td>$(Escape-Html $SysInfo.UserName)</td></tr>
                    <tr><td>Domain:</td><td>$(Escape-Html $SysInfo.Domain)</td></tr>
                    <tr><td>OS:</td><td>$osDisplay</td></tr>
                    <tr><td>Last Reboot:</td><td>$(Escape-Html $lastRebootFormatted)</td></tr>
                    <tr><td>Uptime:</td><td>$(Escape-Html $SysInfo.Uptime)</td></tr>
                </table>
            </td>
            <td colspan="2" style="width: 50%; padding: 4px; vertical-align: top;">
                <table class="section-table" style="width: 100%;">
                    <tr class="section-header">
                        <td colspan="2">HARDWARE</td>
                    </tr>
                    <tr><td>CPU:</td><td>$(Escape-Html $SysInfo.CPU)</td></tr>
"@
    # Add all GPUs
    foreach ($gpu in $SysInfo.GPUs) {
        $gpuLabel = if ($SysInfo.GPUs.Count -gt 1 -and $SysInfo.GPUs.IndexOf($gpu) -gt 0) {
            "GPU #$($SysInfo.GPUs.IndexOf($gpu) + 1):"
        } else {
            "GPU:"
        }
        $html += "                    <tr><td>$gpuLabel</td><td>$(Escape-Html $gpu)</td></tr>`r`n"
    }
    $html += "                    <tr><td>RAM:</td><td>$(Escape-Html $SysInfo.TotalRAM)</td></tr>`r`n"
    $html += "                    <tr><td>Cores:</td><td>$(Escape-Html $SysInfo.Cores)</td></tr>`r`n"
    $html += "                    <tr><td>Connection:</td><td>$(Escape-Html $ConnectionType)</td></tr>`r`n"
    $html += "                    <tr><td>Model:</td><td>$(Escape-Html $SysInfo.SystemModel)</td></tr>`r`n"
    $html += "                </table>`r`n"
    $html += "            </td>`r`n"
    $html += "        </tr>`r`n"
    $html += "        <tr>`r`n"
    $html += "            <td colspan='4' style='padding: 4px;'>`r`n"
    $html += "                <table class='section-table' style='width: 100%; border-collapse: collapse;'>`r`n"
    $html += "                    <tr class='section-header'>`r`n"
    $html += "                        <td colspan='2'>DISK INFORMATION</td>`r`n"
    $html += "                    </tr>`r`n"
    
    # Add disk information as nested table
    if ($SysInfo.Disks.Count -gt 0) {
        $html += "                    <tr><td colspan='2'>`r`n"
        $html += "                        <table class='disk-table' style='width: 100%; table-layout: fixed; border-collapse: collapse;'>`r`n"
        $html += "                            <tr><th>Drive</th><th>Total</th><th>Free</th><th>Usage</th></tr>`r`n"
        foreach ($disk in $SysInfo.Disks) {
            $percentUsed = [math]::Round($disk.PercentUsed, 0)
            $html += "                            <tr>`r`n"
            $html += "                                <td>$(Escape-Html $disk.Label)</td>`r`n"
            $html += "                                <td>$([math]::Round($disk.Total, 1)) GB</td>`r`n"
            $html += "                                <td>$([math]::Round($disk.Free, 1)) GB</td>`r`n"
            $html += "                                <td>$percentUsed%</td>`r`n"
            $html += "                            </tr>`r`n"
        }
        $html += "                        </table>`r`n"
        $html += "                    </td></tr>`r`n"
    } else {
        $html += "                    <tr><td colspan='2'>No disk information available</td></tr>`r`n"
    }
    $html += "                </table>`r`n"
    $html += "            </td>`r`n"
    $html += "        </tr>`r`n"
    $html += "        <tr>`r`n"
    $html += "            <td colspan='4' style='padding: 4px;'>`r`n"
    $html += "                <table class='network-table' style='width: 100%; table-layout: fixed; border-collapse: collapse;'>`r`n"
    
    if ($ActiveAdapters.Count -gt 0) {
        # Check if any adapter has DNS servers
        $hasDNS = ($ActiveAdapters | Where-Object { $_.DNSServers -ne "Unavailable" }).Count -gt 0
        $colspan = if ($hasDNS) { 4 } else { 3 }
        
        $html += "                    <tr class='section-header'>`r`n"
        $html += "                        <td colspan='$colspan'>NETWORK</td>`r`n"
        $html += "                    </tr>`r`n"
        
        # Create horizontal network table with headers
        $html += "                    <tr>`r`n"
        $html += "                        <th style='width: 200px; padding: 3px 6px;'>Adapter</th>`r`n"
        $html += "                        <th style='width: 120px; padding: 3px 6px;'>IP Address</th>`r`n"
        $html += "                        <th style='width: 120px; padding: 3px 6px;'>Gateway</th>`r`n"
        if ($hasDNS) {
            $html += "                        <th style='width: 240px; padding: 3px 6px;'>DNS Servers</th>`r`n"
        }
        $html += "                    </tr>`r`n"
        
        foreach ($adapter in $ActiveAdapters) {
            $html += "                    <tr>`r`n"
            $html += "                        <td style='width: 200px; padding: 3px 6px; word-wrap: break-word; overflow-wrap: break-word;'>$(Escape-Html $adapter.Name)</td>`r`n"
            $html += "                        <td style='width: 120px; padding: 3px 6px;'><span class='ip-address'>$(Escape-Html $adapter.IPAddress)</span></td>`r`n"
            $html += "                        <td style='width: 120px; padding: 3px 6px;'>$(if ($adapter.Gateway) { Escape-Html $adapter.Gateway } else { '-' })</td>`r`n"
            if ($hasDNS) {
                # Format DNS servers to be more compact (replace commas with spaces)
                $dnsDisplay = if ($adapter.DNSServers -ne "Unavailable") {
                    ($adapter.DNSServers -replace ', ', ' ').Trim()
                } else {
                    "-"
                }
                $html += "                        <td style='width: 240px; padding: 3px 6px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;' title='$(Escape-Html $adapter.DNSServers)'>$(Escape-Html $dnsDisplay)</td>`r`n"
            }
            $html += "                    </tr>`r`n"
        }
    } else {
        $html += "                    <tr class='section-header'>`r`n"
        $html += "                        <td colspan='4'>NETWORK</td>`r`n"
        $html += "                    </tr>`r`n"
        $html += "                    <tr><td colspan='4' style='padding: 3px 6px;'>No active network adapters with valid gateway found</td></tr>`r`n"
    }
    $html += "                </table>`r`n"
    $html += "            </td>`r`n"
    $html += "        </tr>`r`n"
    
    $html += "    </table>`r`n"
    $html += "    </div>`r`n"
    $html += "</body>`r`n"
    $html += "</html>`r`n"
    
    return $html
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "System Summary"
$form.Size = New-Object System.Drawing.Size(480, 650)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.BackColor = $backgroundColor
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# Set form icon from GitHub
try {
    $iconUrl = "https://raw.githubusercontent.com/cybredge/public/90dd0143e4a30dd7d1dff5a4d69624133283657d/Assets/logo/icon.ico"
    $iconPath = "$env:TEMP\CybrEdge-Icon.ico"
    
    # Download icon if not already cached
    if (-not (Test-Path $iconPath)) {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($iconUrl, $iconPath)
        $webClient.Dispose()
    }
    
    # Load and set icon
    if (Test-Path $iconPath) {
        $formIcon = New-Object System.Drawing.Icon($iconPath)
        $form.Icon = $formIcon
    }
} catch {
    # Silently fail if icon cannot be loaded
    # Form will use default icon
}

# Root TableLayoutPanel - ONE vertical stack
$rootTable = New-Object System.Windows.Forms.TableLayoutPanel
$rootTable.Dock = [System.Windows.Forms.DockStyle]::Fill
$rootTable.ColumnCount = 1
$rootTable.AutoScroll = $true
$rootTable.Padding = New-Object System.Windows.Forms.Padding(10, 6, 14, 6)
$rootTable.BackColor = $backgroundColor
$rootTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$currentRow = 0

# Helper function to create section
function Add-Section {
    param(
        [string]$Title,
        [System.Windows.Forms.TableLayoutPanel]$DataTable
    )
    
    # Section title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title.ToUpper()
    $titleLabel.Font = $headerFont
    $titleLabel.ForeColor = $headerColor
    $titleLabel.AutoSize = $true
    $rootTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $rootTable.Controls.Add($titleLabel, 0, $currentRow)
    $script:currentRow++
    
    # Thin accent underline
    $divider = New-Object System.Windows.Forms.Panel
    $divider.Height = 1
    $divider.BackColor = $headerColor
    $divider.Margin = New-Object System.Windows.Forms.Padding(0, 1, 0, 3)
    $rootTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 1))) | Out-Null
    $rootTable.Controls.Add($divider, 0, $script:currentRow)
    $script:currentRow++
    
    # Wrap DataTable in a RoundedPanel for border and rounded corners
    $roundedPanel = New-Object RoundedPanel
    $roundedPanel.BackColor = $contentColor
    $roundedPanel.BorderColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $roundedPanel.BorderRadius = 6
    $roundedPanel.BorderWidth = 1
    $roundedPanel.AutoSize = $true
    $roundedPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $roundedPanel.Padding = New-Object System.Windows.Forms.Padding(0)
    
    # Data table - allow natural sizing to prevent wrapping
    $DataTable.Dock = [System.Windows.Forms.DockStyle]::Fill
    $roundedPanel.Controls.Add($DataTable)
    
    $rootTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $rootTable.Controls.Add($roundedPanel, 0, $script:currentRow)
    $script:currentRow++
    
    # Spacing row - increased for better separation
    $rootTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 10))) | Out-Null
    $script:currentRow++
}

# Helper function to create 2-column label/value row
function Add-LabelValueRow {
    param(
        [System.Windows.Forms.TableLayoutPanel]$Table,
        [int]$RowIndex,
        [string]$LabelText,
        [string]$ValueText,
        [System.Drawing.Color]$ValueColor = [System.Drawing.Color]::Black,
        [int]$RowHeight = 20,
        [AllowNull()][System.Drawing.Color]$LabelColorOverride,
        [AllowNull()][System.Drawing.Color]$ValueColorOverride,
        [AllowNull()][System.Drawing.Font]$ValueFontOverride
    )
    
    $Table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $RowHeight))) | Out-Null
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Font = $labelFont
    $label.ForeColor = if ($LabelColorOverride) { $LabelColorOverride } else { $labelColor }
    $label.AutoSize = $true
    $label.AutoEllipsis = $false
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $Table.Controls.Add($label, 0, $RowIndex)
    
    $value = New-Object System.Windows.Forms.Label
    $value.Text = $ValueText
    $value.Font = if ($ValueFontOverride) { $ValueFontOverride } else { $valueFont }
    if ($ValueColorOverride) {
        $value.ForeColor = $ValueColorOverride
    } else {
        $value.ForeColor = if ($ValueColor -eq [System.Drawing.Color]::Black) { $valueColor } else { $ValueColor }
    }
    $value.AutoSize = $true
    $value.AutoEllipsis = $false
    $value.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $value.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $Table.Controls.Add($value, 1, $RowIndex)
}

# ===== DEVICE INFORMATION SECTION =====
$deviceTable = New-Object System.Windows.Forms.TableLayoutPanel
$deviceTable.ColumnCount = 2
$deviceTable.AutoSize = $true
$deviceTable.Dock = [System.Windows.Forms.DockStyle]::Top
$deviceTable.BackColor = $contentColor
$deviceTable.Padding = New-Object System.Windows.Forms.Padding(8, 6, 10, 6)
$deviceTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) | Out-Null
$deviceTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

$deviceRowIndex = 0

# First group: Computer, User, Domain
Add-LabelValueRow -Table $deviceTable -RowIndex $deviceRowIndex -LabelText "Computer Name:" -ValueText $sysInfo.ComputerName
$deviceRowIndex++
Add-LabelValueRow -Table $deviceTable -RowIndex $deviceRowIndex -LabelText "Logged-in User:" -ValueText $sysInfo.UserName
$deviceRowIndex++
Add-LabelValueRow -Table $deviceTable -RowIndex $deviceRowIndex -LabelText "Domain/Workgroup:" -ValueText $sysInfo.Domain
$deviceRowIndex++

    # Small gap above divider
    $deviceTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 3))) | Out-Null
$deviceRowIndex++

# Horizontal divider
$deviceDivider = New-Object System.Windows.Forms.Label
$deviceDivider.Height = 1
$deviceDivider.BackColor = $dividerColor
$deviceDivider.AutoSize = $false
$deviceDivider.Dock = [System.Windows.Forms.DockStyle]::Fill
$deviceTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 1))) | Out-Null
$deviceTable.Controls.Add($deviceDivider, 0, $deviceRowIndex)
$deviceTable.SetColumnSpan($deviceDivider, 2)
$deviceRowIndex++

    # Small gap below divider
    $deviceTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 3))) | Out-Null
$deviceRowIndex++

# Second group: OS (combined Version + Architecture)
$osDisplay = if ($sysInfo.OSArch -ne "Unavailable" -and $sysInfo.OSVersion -ne "Unavailable") {
    "$($sysInfo.OSVersion) ($($sysInfo.OSArch))"
} elseif ($sysInfo.OSVersion -ne "Unavailable") {
    $sysInfo.OSVersion
} else {
    "Unavailable"
}
Add-LabelValueRow -Table $deviceTable -RowIndex $deviceRowIndex -LabelText "OS:" -ValueText $osDisplay
$deviceRowIndex++

# Third group: Last Reboot, Uptime
# Format Last Reboot with 2-digit year and timezone
$timezone = [TimeZoneInfo]::Local
$timezoneAbbr = if ($timezone.Id -match 'Eastern') { if ($timezone.IsDaylightSavingTime($sysInfo.LastReboot)) { 'EDT' } else { 'EST' } }
                elseif ($timezone.Id -match 'Central') { if ($timezone.IsDaylightSavingTime($sysInfo.LastReboot)) { 'CDT' } else { 'CST' } }
                elseif ($timezone.Id -match 'Mountain') { if ($timezone.IsDaylightSavingTime($sysInfo.LastReboot)) { 'MDT' } else { 'MST' } }
                elseif ($timezone.Id -match 'Pacific') { if ($timezone.IsDaylightSavingTime($sysInfo.LastReboot)) { 'PDT' } else { 'PST' } }
                elseif ($timezone.Id -match 'Alaska') { if ($timezone.IsDaylightSavingTime($sysInfo.LastReboot)) { 'AKDT' } else { 'AKST' } }
                elseif ($timezone.Id -match 'Hawaii') { 'HST' }
                else { $timezone.Id }
$lastRebootFormatted = $sysInfo.LastReboot.ToString("M/d/yy h:mm tt") + " " + $timezoneAbbr
Add-LabelValueRow -Table $deviceTable -RowIndex $deviceRowIndex -LabelText "Last Reboot:" -ValueText $lastRebootFormatted
$deviceRowIndex++
Add-LabelValueRow -Table $deviceTable -RowIndex $deviceRowIndex -LabelText "Uptime:" -ValueText $sysInfo.Uptime

Add-Section -Title "Device Information" -DataTable $deviceTable

# ===== HARDWARE SECTION =====
$hardwareTable = New-Object System.Windows.Forms.TableLayoutPanel
$hardwareTable.ColumnCount = 2
$hardwareTable.AutoSize = $true
$hardwareTable.Dock = [System.Windows.Forms.DockStyle]::Top
$hardwareTable.BackColor = $contentColor
$hardwareTable.Padding = New-Object System.Windows.Forms.Padding(10, 8, 14, 8)
$hardwareTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) | Out-Null
$hardwareTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

$hardwareRowIndex = 0

# First group: CPU, GPU(s)
Add-LabelValueRow -Table $hardwareTable -RowIndex $hardwareRowIndex -LabelText "CPU:" -ValueText $sysInfo.CPU
$hardwareRowIndex++

# Add all GPUs
foreach ($gpu in $sysInfo.GPUs) {
    $gpuLabel = if ($sysInfo.GPUs.Count -gt 1 -and $sysInfo.GPUs.IndexOf($gpu) -gt 0) {
        "GPU #$($sysInfo.GPUs.IndexOf($gpu) + 1):"
    } else {
        "GPU:"
    }
    Add-LabelValueRow -Table $hardwareTable -RowIndex $hardwareRowIndex -LabelText $gpuLabel -ValueText $gpu
    $hardwareRowIndex++
}

# Small gap above divider
    $hardwareTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 3))) | Out-Null
$hardwareRowIndex++

# Horizontal divider
$hardwareDivider1 = New-Object System.Windows.Forms.Label
$hardwareDivider1.Height = 1
$hardwareDivider1.BackColor = $dividerColor
$hardwareDivider1.AutoSize = $false
$hardwareDivider1.Dock = [System.Windows.Forms.DockStyle]::Fill
$hardwareTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 1))) | Out-Null
$hardwareTable.Controls.Add($hardwareDivider1, 0, $hardwareRowIndex)
$hardwareTable.SetColumnSpan($hardwareDivider1, 2)
$hardwareRowIndex++

# Small gap below divider
    $hardwareTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 3))) | Out-Null
$hardwareRowIndex++

# Second group: RAM, Cores
Add-LabelValueRow -Table $hardwareTable -RowIndex $hardwareRowIndex -LabelText "RAM:" -ValueText $sysInfo.TotalRAM
$hardwareRowIndex++
Add-LabelValueRow -Table $hardwareTable -RowIndex $hardwareRowIndex -LabelText "# of Cores:" -ValueText $sysInfo.Cores
$hardwareRowIndex++

# Small gap above divider
    $hardwareTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 3))) | Out-Null
$hardwareRowIndex++

# Horizontal divider
$hardwareDivider2 = New-Object System.Windows.Forms.Label
$hardwareDivider2.Height = 1
$hardwareDivider2.BackColor = $dividerColor
$hardwareDivider2.AutoSize = $false
$hardwareDivider2.Dock = [System.Windows.Forms.DockStyle]::Fill
$hardwareTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 1))) | Out-Null
$hardwareTable.Controls.Add($hardwareDivider2, 0, $hardwareRowIndex)
$hardwareTable.SetColumnSpan($hardwareDivider2, 2)
$hardwareRowIndex++

# Small gap below divider
    $hardwareTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 3))) | Out-Null
$hardwareRowIndex++

# Third group: Connection Type, System Model
Add-LabelValueRow -Table $hardwareTable -RowIndex $hardwareRowIndex -LabelText "Connection Type:" -ValueText $connectionType
$hardwareRowIndex++
Add-LabelValueRow -Table $hardwareTable -RowIndex $hardwareRowIndex -LabelText "System Model:" -ValueText $sysInfo.SystemModel

Add-Section -Title "Hardware" -DataTable $hardwareTable

# ===== DISK INFORMATION SECTION =====
$diskTable = New-Object System.Windows.Forms.TableLayoutPanel
$diskTable.ColumnCount = 4
$diskTable.RowCount = $sysInfo.Disks.Count + 2
$diskTable.AutoSize = $true
$diskTable.Dock = [System.Windows.Forms.DockStyle]::Top
$diskTable.BackColor = $contentColor
$diskTable.Padding = New-Object System.Windows.Forms.Padding(8, 6, 10, 6)

$diskTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 50))) | Out-Null
$diskTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null
$diskTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null
$diskTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

# Header row
$diskTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) | Out-Null
$headerLabels = @("Drive", "Total", "Free", "Usage")
for ($i = 0; $i -lt 4; $i++) {
    $header = New-Object System.Windows.Forms.Label
    $header.Text = $headerLabels[$i]
    $header.Font = $labelFont
    $header.ForeColor = $labelColor
    $header.AutoSize = $true
    if ($i -eq 0) {
        $header.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    } else {
        $header.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    }
    $diskTable.Controls.Add($header, $i, 0)
}

# Separator line under header
$separatorRow = New-Object System.Windows.Forms.Label
$separatorRow.Height = 1
$separatorRow.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$separatorRow.AutoSize = $false
$separatorRow.Dock = [System.Windows.Forms.DockStyle]::Fill
$diskTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 1))) | Out-Null
$diskTable.Controls.Add($separatorRow, 0, 1)
$diskTable.SetColumnSpan($separatorRow, 4)

# Data rows
$rowIndex = 2
foreach ($disk in $sysInfo.Disks) {
    $diskTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26))) | Out-Null
    
    # Drive (left-aligned)
    $driveLabel = New-Object System.Windows.Forms.Label
    $driveLabel.Text = $disk.Label
    $driveLabel.Font = $valueFont
    $driveLabel.ForeColor = $valueColor
    $driveLabel.AutoSize = $true
    $driveLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $diskTable.Controls.Add($driveLabel, 0, $rowIndex)
    
    # Total (right-aligned) - swapped position
    $totalLabel = New-Object System.Windows.Forms.Label
    $totalLabel.Text = "{0:N1} GB" -f $disk.Total
    $totalLabel.Font = $valueFont
    $totalLabel.ForeColor = $valueColor
    $totalLabel.AutoSize = $true
    $totalLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $diskTable.Controls.Add($totalLabel, 1, $rowIndex)
    
    # Free (right-aligned) - swapped position
    $freeLabel = New-Object System.Windows.Forms.Label
    $freeLabel.Text = "{0:N1} GB" -f $disk.Free
    $freeLabel.Font = $valueFont
    $freeLabel.ForeColor = $valueColor
    $freeLabel.AutoSize = $true
    $freeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $diskTable.Controls.Add($freeLabel, 2, $rowIndex)
    
    # Usage Bar container
    $barContainer = New-Object System.Windows.Forms.TableLayoutPanel
    $barContainer.ColumnCount = 2
    $barContainer.RowCount = 1
    $barContainer.AutoSize = $true
    $barContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $barContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 50))) | Out-Null
    $barContainer.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)
    $barContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) | Out-Null
    
    $percentUsed = [math]::Round($disk.PercentUsed, 0)
    $barColor = if ($percentUsed -ge 85) { [System.Drawing.Color]::FromArgb(220, 53, 69) } elseif ($percentUsed -ge 70) { [System.Drawing.Color]::FromArgb(255, 193, 7) } else { [System.Drawing.Color]::FromArgb(40, 167, 69) }
    
    # Background track panel (faint background)
    $trackPanel = New-Object System.Windows.Forms.Panel
    $trackPanel.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)  # Faint gray background track
    $trackPanel.Height = 14
    $trackPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $trackPanel.Margin = New-Object System.Windows.Forms.Padding(0)
    
    # Progress bar (kept exactly as-is)
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = $percentUsed
    $progressBar.Height = 16
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.ForeColor = $barColor
    $progressBar.Dock = [System.Windows.Forms.DockStyle]::Fill
    $trackPanel.Controls.Add($progressBar)
    
    $barContainer.Controls.Add($trackPanel, 0, 0)
    
    # Percentage label (vertically centered with bar)
    $percentLabel = New-Object System.Windows.Forms.Label
    $percentLabel.Text = "$percentUsed%"
    $percentLabel.Font = $labelFont
    $percentLabel.ForeColor = $valueColor
    $percentLabel.AutoSize = $true
    $percentLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $percentLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $barContainer.Controls.Add($percentLabel, 1, 0)
    
    $diskTable.Controls.Add($barContainer, 3, $rowIndex)
    
    $rowIndex++
}

Add-Section -Title "Disk Information" -DataTable $diskTable

# ===== NETWORK SECTION =====
$networkTable = New-Object System.Windows.Forms.TableLayoutPanel
$networkTable.ColumnCount = 2
$networkTable.AutoSize = $true
$networkTable.Dock = [System.Windows.Forms.DockStyle]::Top
$networkTable.BackColor = $contentColor
$networkTable.Padding = New-Object System.Windows.Forms.Padding(8, 6, 10, 6)
$networkTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) | Out-Null
$networkTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

$netRowIndex = 0

if ($activeAdapters.Count -gt 0) {
    $adapterIndex = 0
    foreach ($adapter in $activeAdapters) {
        # Add divider and spacing before each adapter except the first
        if ($adapterIndex -gt 0) {
            # Top margin before new adapter block
            $networkTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 5))) | Out-Null
            $netRowIndex++
            
            # Horizontal divider
            $divider = New-Object System.Windows.Forms.Label
            $divider.Height = 1
            $divider.BackColor = $dividerColor
            $divider.AutoSize = $false
            $divider.Dock = [System.Windows.Forms.DockStyle]::Fill
            $networkTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 1))) | Out-Null
            $networkTable.Controls.Add($divider, 0, $netRowIndex)
            $networkTable.SetColumnSpan($divider, 2)
            $netRowIndex++
            
            # Small gap below divider
            $networkTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 2))) | Out-Null
            $netRowIndex++
        }
        
        # Adapter name (slightly bold for visibility)
        $adapterNameFont = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        Add-LabelValueRow -Table $networkTable -RowIndex $netRowIndex -LabelText "Adapter:" -ValueText $adapter.Name -ValueFontOverride $adapterNameFont -RowHeight 20
        $netRowIndex++
        
        # IP Address (highlighted)
        Add-LabelValueRow -Table $networkTable -RowIndex $netRowIndex -LabelText "IP Address:" -ValueText $adapter.IPAddress -ValueColor $accentColor -RowHeight 20
        $netRowIndex++
        
        # Gateway (reduced row height)
        Add-LabelValueRow -Table $networkTable -RowIndex $netRowIndex -LabelText "Gateway:" -ValueText $adapter.Gateway -RowHeight 20
        $netRowIndex++
        
        # Subnet (if available, reduced row height)
        if ($adapter.Subnet) {
            Add-LabelValueRow -Table $networkTable -RowIndex $netRowIndex -LabelText "Subnet:" -ValueText $adapter.Subnet -RowHeight 20
            $netRowIndex++
        }
        
        # DNS Servers (if available, reduced row height)
        if ($adapter.DNSServers -ne "Unavailable") {
            Add-LabelValueRow -Table $networkTable -RowIndex $netRowIndex -LabelText "DNS Servers:" -ValueText $adapter.DNSServers -RowHeight 20
            $netRowIndex++
        }
        
        $adapterIndex++
    }
} else {
    # No active adapters with gateways found
    Add-LabelValueRow -Table $networkTable -RowIndex $netRowIndex -LabelText "Status:" -ValueText "No active network adapters with valid gateway found"
}

Add-Section -Title "Network" -DataTable $networkTable

# Footer panel with buttons
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$footerPanel.Height = 45
$footerPanel.BackColor = $backgroundColor
$footerPanel.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 6)

# Copy button
$copyButtonColor = [System.Drawing.Color]::White
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "Copy"
$copyButton.Size = New-Object System.Drawing.Size(100, 26)
$copyButton.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
$copyButton.BackColor = $copyButtonColor  # White button
$copyButton.ForeColor = [System.Drawing.Color]::Black  # Black text
$copyButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$copyButton.FlatAppearance.BorderSize = 1
$copyButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)  # Light gray border
$copyButton.Cursor = [System.Windows.Forms.Cursors]::Hand

# Copy button click handler
$copyButton.Add_Click({
    try {
        $formattedHtml = Format-SystemInfoForEmail -SysInfo $sysInfo -ConnectionType $connectionType -ActiveAdapters $activeAdapters
        Set-ClipboardHtml -html $formattedHtml
        
        # Show brief confirmation
        $copyButton.Text = "Copied!"
        $copyButton.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)  # Green
        $copyButton.ForeColor = [System.Drawing.Color]::White
        $form.Refresh()
        
        # Reset button after 1.5 seconds
        $script:timer = New-Object System.Windows.Forms.Timer
        $script:timer.Interval = 1500
        $script:timer.Add_Tick({
            $copyButton.Text = "Copy"
            $copyButton.BackColor = $copyButtonColor
            $copyButton.ForeColor = [System.Drawing.Color]::Black
            $script:timer.Stop()
            $script:timer.Dispose()
            $script:timer = $null
        })
        $script:timer.Start()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Exit button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Close"
$exitButton.Size = New-Object System.Drawing.Size(100, 26)
$exitButton.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
$exitButton.BackColor = $headerColor
$exitButton.ForeColor = [System.Drawing.Color]::White
$exitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$exitButton.FlatAppearance.BorderSize = 1
$exitButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)  # Light gray border
$exitButton.Cursor = [System.Windows.Forms.Cursors]::Hand

# Exit button click handler
$exitButton.Add_Click({
    $form.Close()
})

# Center buttons using TableLayoutPanel
$footerTable = New-Object System.Windows.Forms.TableLayoutPanel
$footerTable.Dock = [System.Windows.Forms.DockStyle]::Fill
$footerTable.ColumnCount = 5
$footerTable.RowCount = 1
$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10))) | Out-Null
$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$footerTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$footerTable.Controls.Add($copyButton, 1, 0)
$footerTable.Controls.Add($exitButton, 3, 0)

$footerPanel.Controls.Add($footerTable)

# Add form controls
$form.Controls.Add($rootTable)
$form.Controls.Add($footerPanel)

# Calculate required width based on content and add buffer
# Measure the widest content to ensure nothing wraps
$maxWidth = 0
$graphics = $form.CreateGraphics()

# Check Device Information table
foreach ($control in $deviceTable.Controls) {
    if ($control -is [System.Windows.Forms.Label]) {
        $textSize = $graphics.MeasureString($control.Text, $control.Font)
        $controlWidth = $textSize.Width + $control.Margin.Left + $control.Margin.Right
        if ($controlWidth -gt $maxWidth) { $maxWidth = $controlWidth }
    }
}

# Check Hardware table
foreach ($control in $hardwareTable.Controls) {
    if ($control -is [System.Windows.Forms.Label]) {
        $textSize = $graphics.MeasureString($control.Text, $control.Font)
        $controlWidth = $textSize.Width + $control.Margin.Left + $control.Margin.Right
        if ($controlWidth -gt $maxWidth) { $maxWidth = $controlWidth }
    }
}

# Check Network table
foreach ($control in $networkTable.Controls) {
    if ($control -is [System.Windows.Forms.Label]) {
        $textSize = $graphics.MeasureString($control.Text, $control.Font)
        $controlWidth = $textSize.Width + $control.Margin.Left + $control.Margin.Right
        if ($controlWidth -gt $maxWidth) { $maxWidth = $controlWidth }
    }
}

$graphics.Dispose()

# Calculate total width needed: label column (130) + max content + padding (12+18) + root padding (16+24) + buffer
$requiredWidth = 130 + $maxWidth + 12 + 18 + 16 + 24 + 40
if ($requiredWidth -lt 480) { $requiredWidth = 480 }
$form.Width = $requiredWidth

# Calculate required height to show Device Info, Hardware, Disk Info, and first Network adapter
$requiredHeight = 0

# Root padding top
$requiredHeight += 6

# Device Information section
$requiredHeight += 18  # Title height (AutoSize, reduced font)
$requiredHeight += 1   # Divider
$requiredHeight += 1   # Divider margin top (reduced)
$requiredHeight += 3   # Divider margin bottom (reduced)
$deviceRowCount = 6  # Computer, User, Domain, OS, Last Reboot, Uptime
$requiredHeight += ($deviceRowCount * 20)  # Rows (reduced from 22)
$requiredHeight += 12  # Table padding (top + bottom: 6+6, reduced)
$requiredHeight += 8   # Section spacing (reduced)

# Hardware section
$requiredHeight += 18  # Title height (reduced font)
$requiredHeight += 1   # Divider
$requiredHeight += 1   # Divider margin top (reduced)
$requiredHeight += 3   # Divider margin bottom (reduced)
$hardwareRowCount = 1  # CPU
$hardwareRowCount += $sysInfo.GPUs.Count  # GPU(s)
$hardwareRowCount += 2  # RAM, Cores
$hardwareRowCount += 2  # Connection Type, System Model
$requiredHeight += ($hardwareRowCount * 20)  # Regular rows (reduced from 22)
$requiredHeight += 5   # First divider section (reduced)
$requiredHeight += 5   # Second divider section (reduced)
$requiredHeight += 12  # Table padding (top + bottom: 6+6, reduced)
$requiredHeight += 8   # Section spacing (reduced)

# Disk Information section
$requiredHeight += 18  # Title height (reduced font)
$requiredHeight += 1   # Divider
$requiredHeight += 1   # Divider margin top (reduced)
$requiredHeight += 3   # Divider margin bottom (reduced)
$requiredHeight += 18  # Header row (reduced font)
$requiredHeight += 1   # Separator line
$diskRowCount = if ($sysInfo.Disks.Count -gt 0) { $sysInfo.Disks.Count } else { 1 }
$requiredHeight += ($diskRowCount * 24)  # Data rows (reduced from 26)
$requiredHeight += 12  # Table padding (top + bottom: 6+6, reduced)
$requiredHeight += 8   # Section spacing (reduced)

# Network section - first adapter only
$requiredHeight += 18  # Title height (reduced font)
$requiredHeight += 1   # Divider
$requiredHeight += 1   # Divider margin top (reduced)
$requiredHeight += 3   # Divider margin bottom (reduced)
if ($activeAdapters.Count -gt 0) {
    # First adapter has NO top margin (only subsequent adapters do)
    # Count all rows for first adapter: Adapter name, IP Address, Gateway (always shown), Subnet (if exists), DNS Servers (always shown)
    # Actual rows: 5 (Adapter, IP, Gateway, Subnet, DNS)
    # However, AutoSize elements add extra height:
    # - Title label with AutoSize: ~22px actual (9pt Bold font) vs estimated 18px = +4px
    # - RoundedPanel wrapper: border (1px top + 1px bottom) + border radius overhead = ~4px
    # - TableLayoutPanel AutoSize spacing: ~2px
    # - Safety buffer for different DPI/font rendering: ~30px
    # Total extra needed: ~40px = 2 rows worth
    # So we use 7 rows (5 actual + 2 buffer) to ensure all fields are visible across all systems
    $firstAdapterRows = 5  # Actual data rows: Adapter, IP, Gateway, Subnet, DNS
    $firstAdapterRows += 2  # Buffer for AutoSize elements, RoundedPanel wrapper, and rendering differences
    $requiredHeight += ($firstAdapterRows * 20)  # All rows use RowHeight 20
    $requiredHeight += 12  # Table padding (top + bottom: 6+6, reduced)
} else {
    # Even if no adapters, add some height for the "No adapters" message
    $requiredHeight += 20  # Minimum height for empty state
    $requiredHeight += 12  # Table padding
}

# Footer
$requiredHeight += 45  # Reduced from 50

# Root padding bottom
$requiredHeight += 6

# Add buffer for form borders, title bar, etc. (increased to ensure all content is visible)
$requiredHeight += 50

# Ensure minimum height to display Device Info, Hardware, Disk Info, Network section, and ALL fields of first adapter
# With reduced sizes, calculate a more accurate minimum that includes all adapter fields
# Network section: Title(18) + Divider(1) + DividerMarginTop(1) + DividerMarginBottom(3) + AdapterRows(5*20) + TablePadding(12) = 18+1+1+3+100+12 = 135
# Note: First adapter has NO top margin (only subsequent adapters do)
# Always use maximum rows (5) to ensure Subnet and DNS are visible even if Subnet doesn't exist
$minHeight = 6 + 18 + 1 + 1 + 3 + (6 * 20) + 12 + 8 +  # Device Info: RootPaddingTop(6) + Title(18) + Divider(1) + MarginTop(1) + MarginBottom(3) + Rows(6*20) + TablePadding(12) + SectionSpacing(8)
             18 + 1 + 1 + 3 + (6 * 20) + 5 + 5 + 12 + 8 +  # Hardware: Title(18) + Divider(1) + MarginTop(1) + MarginBottom(3) + Rows(6*20) + Dividers(5+5) + TablePadding(12) + SectionSpacing(8)
             18 + 1 + 1 + 3 + 18 + 1 + (1 * 24) + 12 + 8 +  # Disk Info: Title(18) + Divider(1) + MarginTop(1) + MarginBottom(3) + HeaderRow(18) + Separator(1) + DataRows(1*24) + TablePadding(12) + SectionSpacing(8)
             18 + 1 + 1 + 3 + (7 * 20) + 12 +  # Network: Title(18) + Divider(1) + MarginTop(1) + MarginBottom(3) + AdapterRows(7*20: 5 actual + 2 buffer for AutoSize/RoundedPanel) + TablePadding(12) - NO top margin for first adapter
             45 + 6 + 50  # Footer(45) + RootPaddingBottom(6) + Buffer(50)
# Ensure height is at least enough to show all first adapter fields (always use minimum that includes all 5 rows)
if ($requiredHeight -lt $minHeight) { $requiredHeight = $minHeight }
$form.Height = $requiredHeight

# Resize handler to keep tables full-width (but allow AutoSize columns to expand)
$rootTable.Add_Resize({
    $rootWidth = $rootTable.ClientSize.Width
    $diskTable.Width = $rootWidth
    # Device, Hardware, and Network tables use AutoSize for value column, so they'll expand naturally
})

# Show form
$form.ShowDialog()
