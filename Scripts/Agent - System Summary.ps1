Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function FormatHeaderLine {
    param (
        [string]$label,
        [int]$width
    )
    $label = " $($label.ToUpper()) "
    $sideLength = [Math]::Floor(($width - $label.Length) / 2)
    $left = "=" * $sideLength
    $right = "=" * ($width - $sideLength - $label.Length)
    return "$left$label$right"
}

function Get-SystemInfoText {
    $nl = "`r`n"
    $text = ""

    # Device Info
    $computerName = [string]$env:COMPUTERNAME
    $userName     = [string]$env:USERNAME
    $compSys      = Get-WmiObject Win32_ComputerSystem
    $domain       = [string]$compSys.Domain

    # OS Info
    $os           = Get-WmiObject Win32_OperatingSystem
    $osVersion    = [string]$os.Caption
    $osArch       = if ($os.OSArchitecture) { $os.OSArchitecture } else { "Unavailable" }
    $lastReboot   = $os.ConvertToDateTime($os.LastBootUpTime)
    $uptimeSpan   = (Get-Date) - $lastReboot
    $uptime       = "{0} days {1} hours" -f $uptimeSpan.Days, [Math]::Floor($uptimeSpan.Hours)

    # Hardware Info
    $totalRAMGB   = if ($compSys) { [math]::Round($compSys.TotalPhysicalMemory / 1GB, 1).ToString() + " GB" } else { "Unavailable" }

    $cpuObj       = Get-WmiObject Win32_Processor | Select-Object -First 1
    $cpu          = if ($cpuObj) { $cpuObj.Name.Trim() } else { "Unavailable" }

    $gpuObj       = Get-WmiObject Win32_VideoController | Select-Object -First 1
    $gpu          = if ($gpuObj) { $gpuObj.Name } else { "Unavailable" }

    $manufacturer = if ($compSys.Manufacturer) { $compSys.Manufacturer.Trim() } else { "" }
    $model        = if ($compSys.Model -and $compSys.Model -notmatch "System Product Name") { $compSys.Model.Trim() } else { "" }
    $sysModel     = "$manufacturer $model".Trim()
    if (-not $sysModel) { $sysModel = "Unavailable" }

    # Disk Info
    $disks = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
        $label = $_.DeviceID
        $free  = [math]::Round($_.FreeSpace / 1GB, 1)
        $total = [math]::Round($_.Size / 1GB, 1)
        "{0,-4} Free: {1,7} GB | Total: {2,7} GB" -f $label, $free, $total
    }

    # Network Info
    $adapter = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled } | Select-Object -First 1
    if ($adapter) {
        $adapterName = [string]$adapter.Description
        $ipAddress   = [string]($adapter.IPAddress | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
        $subnet      = [string]($adapter.IPSubnet | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
        $gateway     = [string]($adapter.DefaultIPGateway | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
        $dnsServers  = ($adapter.DNSServerSearchOrder | Where-Object { $_ -match '^\d+\.' }) -join ", "
    } else {
        $adapterName = $ipAddress = $subnet = $gateway = $dnsServers = "Unavailable"
    }

    # Width estimation
    $tempText = @()
    $tempText += FormatHeaderLine "Device Info" 50
    $tempText += "Computer Name     : $computerName"
    $tempText += "Logged In User    : $userName"
    $tempText += "Domain/Group      : $domain"
    $tempText += ""
    $tempText += FormatHeaderLine "Operating System" 50
    $tempText += "OS Version        : $osVersion"
    $tempText += "Architecture      : $osArch"
    $tempText += "Last Reboot       : $($lastReboot.ToString("g"))"
    $tempText += "Uptime            : $uptime"
    $tempText += ""
    $tempText += FormatHeaderLine "Hardware" 50
    $tempText += "Installed RAM     : $totalRAMGB"
    $tempText += "CPU               : $cpu"
    $tempText += "GPU               : $gpu"
    $tempText += "System Model      : $sysModel"
    $tempText += ""
    $tempText += FormatHeaderLine "Disks" 50
    $tempText += $disks
    $tempText += ""
    $tempText += FormatHeaderLine "Network" 50
    $tempText += "Adapter Name      : $adapterName"
    $tempText += "IP Address        : $ipAddress"
    $tempText += "Subnet Mask       : $subnet"
    $tempText += "Gateway           : $gateway"
    $tempText += "DNS Servers       : $dnsServers"

    $longest = ($tempText | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum

    # Final formatted output
    $text += $(FormatHeaderLine "Device Info" $longest) + $nl
    $text += "Computer Name     : $computerName$nl"
    $text += "Logged In User    : $userName$nl"
    $text += "Domain/Group      : $domain$nl$nl"

    $text += $(FormatHeaderLine "Operating System" $longest) + $nl
    $text += "OS Version        : $osVersion$nl"
    $text += "Architecture      : $osArch$nl"
    $text += "Last Reboot       : $($lastReboot.ToString("g"))$nl"
    $text += "Uptime            : $uptime$nl$nl"

    $text += $(FormatHeaderLine "Hardware" $longest) + $nl
    $text += "Installed RAM     : $totalRAMGB$nl"
    $text += "CPU               : $cpu$nl"
    $text += "GPU               : $gpu$nl"
    $text += "System Model      : $sysModel$nl$nl"

    $text += $(FormatHeaderLine "Disks" $longest) + $nl
    $text += ($disks -join $nl) + $nl + $nl

    $text += $(FormatHeaderLine "Network" $longest) + $nl
    $text += "Adapter Name      : $adapterName$nl"
    $text += "IP Address        : $ipAddress$nl"
    $text += "Subnet Mask       : $subnet$nl"
    $text += "Gateway           : $gateway$nl"
    $text += "DNS Servers       : $dnsServers"

    return ,$text.TrimEnd(), $longest
}

# Build the Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "System Summary"
$form.StartPosition = "CenterScreen"
$form.Topmost = $true
$form.AutoSize = $true
$form.AutoSizeMode = "GrowAndShrink"
$form.FormBorderStyle = "FixedDialog"

# Get content
$result = Get-SystemInfoText
$textContent = $result[0]
$longest = $result[1]

$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Multiline = $true
$textbox.ReadOnly = $true
$textbox.Font = New-Object System.Drawing.Font("Consolas", 10)
$textbox.Text = $textContent
$textbox.BorderStyle = "None"
$textbox.AutoSize = $true

# Size calculation
$lines = $textbox.Lines.Count
$charWidth = 7
$charHeight = 18
$textbox.Width = ($longest + 2) * $charWidth
$textbox.Height = $lines * $charHeight

$form.Controls.Add($textbox)
$form.ClientSize = $textbox.Size
$form.ShowDialog()
