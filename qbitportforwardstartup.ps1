# First, right-click and open up notepad as administrator, paste the entire script, and save it as DESIRED_NAME.ps1 so that it gets saved as a Powershell script.
# You can put the Powershell script in the same folder as your qBit exe file if you want.
# This Powershell script is meant to be used with windows task scheduler, with the following trigger. Create new task, not a basic task.
# Trigger: "On an event"
# Log: Microsoft-Windows-NetworkProfile/Operation
# Source: NetworkProfile
# Event ID: 10000

# Also be sure to check "Start only if the following network connection is available" in the Conditions tab then select "wgpia0" which is the network connection for PIA VPN WireGuard.
# If you want to use OpenVPN Protocol for PIA you might need another method to modify the trigger, as OpenVPN does not create a network name you can just point to.
# This trigger ensures that the actions you want to run only start when the PIA network connection has been fully established.
# If you want to autorun qBit at startup, you want this trigger because qBit does not automatically re-establish a connection if you open it up before a connection is made.
# You may need to select "Run with highest privileges" in the General tab of the task window. You may also need to change the "Configure for" setting to Windows 10.
# You might want to uncheck "Stop the task if it runs longer than:" in the Settings tab.

# You may need to change the security options so that your "Administrators" account triggers the task instead of your user account.
# To do this, click on "Change User or Group..." in the General tab of the task window.
# Click on "Advanced..." in the new window. Then click on "Find Now" on the right side of the newer window.
# Click on either "Administrators" or your microsoft username that has your email attached to it. Note: If you want to use "Administrators," make sure you are not clicking on "Administrator."
# Press OK.

# In the "Actions" tab of the task window, click on "New..." and select "Start a program" as your action. Then point to the Powershell.exe, not the script or qBit.
# Powershell is most likely located in "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
# If you're not sure where it is, press the windows button and type Powershell, then right-click Powershell and click on open file location.
# This will lead you to the shortcut for Powershell, right-click the shortcut then open file location again and you should see the Powershell exe.
# In the "Add arguments" section of the same window, you want to paste the full filepath of your script and type -File before it like below.
# -File "C:\Program Files\qBittorrent\qbitportforwardstartup.ps1"
# Press OK

# Note: This task will also cause qBit to close when your VPN disconnects, but not if you close PIA yourself. It will also reopen qBit whenever your VPN connects again, not just at startup.
# The script along with the task should ensure that the port forward number for PIA is checked & appropriately changed in qBit every time (Right before) your qBit starts on its own through the task trigger/script.

# Ensure that Powershell is in unrestricted mode so that you can execute scripts with it. Right-click on Powershell and open as administrator. Type the below in Powershell window.
# Set-ExecutionPolicy unrestricted
# Type "Get-ExecutionPolicy" in the same window to check that the policy is now unrestricted.

# Replace "YOUR USER DIRECTORY" in line 53 of this script with your windows account name. 
# If you're not sure, check the name by going into "Users" folder in your main hard drive & click on the white address bar at the top of your window to reveal the file path.

# Actual start of script:
# X second lag to let PIA VPN retrieve port forward number after making a full connection. Adjust the amount to fit with your setup.
# The moment when PIA interface reveals the port forward number may be a few moments after the moment when the port forward number is actually retrieved.
Start-Sleep -Seconds 5

Write-Host "Checking for forwarded port."
# Check the forwarded port by PIA
$PortPia = & "C:\Program Files\Private Internet Access\piactl.exe" get portforward
# Convert variable to integer
$PortPia = $PortPia -as [int]

# Replace line in the .ini file
Write-Host "Setting port to" $PortPia"."
$FichierIni = "C:\Users\YOUR USER DIRECTORY\AppData\Roaming\qBittorrent\qBittorrent.ini"
(Get-Content $FichierIni) -replace "Connection\\PortRangeMin=\d*", "Connection\PortRangeMin=$PortPia" | Set-Content -Path $FichierIni

# Make sure the following is the actual path to your qBit exe file, or change the line to where you have it saved.
Write-Host "Starting up qBit."
Start-Process -FilePath "C:\Program Files\qBittorrent\qbittorrent.exe"
