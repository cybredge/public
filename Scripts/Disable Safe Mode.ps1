<# 

Purpose: PowerShell Script to Revert Windows back int0 Normal Boot Mode and Restart
- Meant to be saved as a start up script to revert the safe mode reboot
- Make sure this script is placed in the public repo in GitHub

#>





# Check if the script is running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))

{

    Write-Warning "Please run this script as an Administrator!"

    Break
    
}





# Run command to boot into normal mode

Try {


    # Disable safe boot
    
    cmd.exe /c "bcdedit /deletevalue {current} safeboot"


    # Wait 5 seconds then restart the computer without prompt


    Start-Sleep 5


    Restart-Computer -Force


} Catch {


    Write-Error "Failed to set Normal Mode. Please ensure you are running this as Administrator."


}