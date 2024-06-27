<#

.DESCRIPTION
    This script creates a new Edge profile with the specified name and settings.

.PARAMETER Name
    This parameter is a string and is mandatory.
    The name of the new Edge profile.

.PARAMETER Image
    This parameter is a string and is optional.
    The URL of the image to use as the icon for the new Edge profile.

.PARAMETER StartMenuShortcut
    This parameter is a switch and is optional.
    Creates a Start Menu shortcut for the new Edge profile.

.EXAMPLE
    Create-EdgeProfile -Name "Example"
    This example creates a new Edge profile named "Example".

.EXAMPLE
    Create-EdgeProfile -Name "Example" -Image "https://example.com/image.png"
    This example creates a new Edge profile named "Example" with the specified image as the icon.

.EXAMPLE
    Create-EdgeProfile -Name "Example" -Image "https://example.com/image.png" -StartMenuShortcut
    This example creates a new Edge profile named "Example" with the specified image as the icon and creates a Start Menu shortcut for the profile.

.EXAMPLE
    Create-EdgeProfile -Name "Example" -StartMenuShortcut
    This example creates a new Edge profile named "Example" and creates a Start Menu shortcut for the profile.

.NOTES
    1.0 - 2024-06-27 - Initial version.

#>

# Parameters
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[a-zA-Z0-9]+$")]
    [string]$Name,

    [Parameter(Mandatory = $false)]
    [string]$Image,

    [Parameter(Mandatory = $false)]
    [switch]$StartMenuShortcut
)

# Variables
$ProfileFolder = "Profile-" + $Name
$ProfilePath = "$Env:LOCALAPPDATA\Microsoft\Edge\User Data\$ProfileFolder"
$EdgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# Check if script is running as administrator
if ($StartMenuShortcut) {
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You need to run this script as an administrator to create a Start Menu shortcut."
        Break
    }
}

# Check if profile already exists
if ((Test-Path $ProfilePath) -eq $true) {
    Write-Error "An Edge profile already exists with this name. Try another one."
    Break
}

# Create profile
try {
    Start-Process -FilePath $EdgePath -ArgumentList "--profile-directory=$ProfileFolder --no-first-run --no-default-browser-check --flag-switches-begin --flag-switches-end --site-per-process" | Out-Null
}
catch {
    Write-Error "Failed to start Edge, make sure you have Edge installed in the default location."
    Break
}

# Wait for Edge to create the profile
Write-Output "Waiting for Edge to create the profile..."
Start-Sleep -Seconds 15

# Close Edge
try {
    Stop-Process -Name "msedge"
}
catch {
    Write-Error "Failed to close Edge. Make sure to not close Edge manually while the script is running."
    Break
}

# Backup Local State
try {
    $LocalStateFile = "$Env:LOCALAPPDATA\Microsoft\Edge\User Data\Local State"
    $LocalStateBackUp = "$Env:LOCALAPPDATA\Microsoft\Edge\User Data\Local State Backup"
    Copy-Item $LocalStateFile -Destination $LocalStateBackUp
}
catch {
    Write-Error "Failed to backup Local State. Make sure that `"$LocalStateFile`" exists."
    Break
}

# Write profile name to Local State
try {
    $State = Get-Content -Raw $LocalStateFile
    $Json = $State | ConvertFrom-Json
    $NewEdgeProfile = $Json.profile.info_cache.$ProfileFolder
    $NewEdgeProfile.name = $Name
    $Json | ConvertTo-Json -Compress -Depth 100 | Out-File $LocalStateFile -Encoding UTF8
}
catch {
    Write-Error "Failed to write profile name to Local State."
    Break
}

# Write profile name to registry
try {
    Push-Location
    Set-Location HKCU:\Software\Microsoft\Edge\Profiles\$ProfileFolder
    Set-ItemProperty . ShortcutName "$Name"
    Pop-Location
}
catch {
    Write-Error "Failed to write profile name to registry."
    Break
}

# Backup Preferences
try {
    $PreferencesSettings = "$ProfilePath\Preferences"
    $PreferencesSettingsBackup = "$ProfilePath\Preferences Backup"
    Copy-Item $PreferencesSettings -Destination $PreferencesSettingsBackup
}
catch {
    Write-Error "Failed to backup Preferences. Make sure that `"$ProfilePath\Preferences`" exists."
    Break
}

# Load Preferences
try {
    $Preferencess = Get-Content -Raw $PreferencesSettings
    $PreferencesJson = $Preferencess | ConvertFrom-Json
}
catch {
    Write-Error "Failed to load Preferences. Make sure that `"$ProfilePath\Preferences`" exists."
    Break
}

# Disable Sidebar
if ($null -eq $PreferencesJson.browser.show_hub_apps_tower) {
    try {
        $PreferencesJson.browser | add-member -Name "show_hub_apps_tower" -value $false -MemberType NoteProperty
    }
    catch {
        Write-Error "Failed to add show_hub_apps_tower to Preferences."
        Break
    }
}
else {
    try {
        $PreferencesJson.browser.show_hub_apps_tower = $false
    }
    catch {
        Write-Error "Failed to set show_hub_apps_tower to false in Preferences."
        Break
    }
}
if ($null -eq $PreferencesJson.browser.show_hub_apps_tower_pinned) {
    try {
        $PreferencesJson.browser | add-member -Name "show_hub_apps_tower_pinned" -value $false -MemberType NoteProperty
    }
    catch {
        Write-Error "Failed to add show_hub_apps_tower_pinned to Preferences."
        Break
    }
}
else {
    try {
        $PreferencesJson.browser.show_hub_apps_tower_pinned = $false
    }
    catch {
        Write-Error "Failed to set show_hub_apps_tower_pinned to false in Preferences."
        Break
    }
}

if ($null -eq $PreferencesJson.browser.show_toolbar_learning_toolkit_button) {
    try {
        $PreferencesJson.browser | add-member -Name "show_toolbar_learning_toolkit_button" -value $false -MemberType NoteProperty
    }
    catch {
        Write-Error "Failed to add show_toolbar_learning_toolkit_button to Preferences."
        Break
    }
}
else {
    try {
        $PreferencesJson.browser.show_toolbar_learning_toolkit_button = $false
    }
    catch {
        Write-Error "Failed to set show_toolbar_learning_toolkit_button to false in Preferences."
        Break
    }
}

# Disable data share between profiles
if ($null -eq $PreferencesJson.local_browser_data_share.enabled) {
    $Value = @"
{
    "enabled": false,
    "index_last_cleaned_time": "0"
}
"@
    try {
        $PreferencesJson | add-member -Name "local_browser_data_share" -value (Convertfrom-Json $Value) -MemberType NoteProperty
    }
    catch {
        Write-Error "Failed to add local_browser_data_share to Preferences."
        Break
    }
}
else {
    try {
        $PreferencesJson.local_browser_data_share.enabled = $false
    }
    catch {
        Write-Error "Failed to disable local_browser_data_share in Preferences."
        Break
    }
}

# Disable account based profile switching
if ($null -eq $PreferencesJson.guided_switch.enabled) {
    $Value = @"
{
    "enabled": false
}
"@
    try {
        $PreferencesJson | add-member -Name "guided_switch" -value (Convertfrom-Json $Value) -MemberType NoteProperty
    }
    catch {
        Write-Error "Failed to add guided_switch to Preferences."
        Break
    }
}
else {
    try {
        $PreferencesJson.guided_switch.enabled = $false
    }
    catch {
        Write-Error "Failed to disable guided_switch in Preferences."
        Break
    }
}

# Edit start page settings
$Value = @"
{
    "background_image_type":"imageAndVideo",
    "hide_default_top_sites":false,
    "layout_mode":3,
    "news_feed_display":"off",
    "num_personal_suggestions":1,
    "prerender_contents_height":823,
    "prerender_contents_width":1185,
    "quick_links_options":0
}
"@
if ($null -eq $PreferencesJson.ntp) {
    try {
        $PreferencesJson | add-member -Name "ntp" -value (Convertfrom-Json $Value) -MemberType NoteProperty
    }
    catch {
        Write-Error "Failed to add ntp to Preferences."
        Break
    }
}
else {
    try {
        $PreferencesJson | add-member -Name "ntp" -value (Convertfrom-Json $Value) -MemberType NoteProperty -Force
    }
    catch {
        Write-Error "Failed to set ntp in Preferences."
        Break
    }
}

# Write new settings to preferences
$PreferencesJson | ConvertTo-Json -Compress -Depth 100 | Out-File $PreferencesSettings -Encoding UTF8

if ($Image -ne "") {
    # Download image
    try {
        Invoke-WebRequest -Uri $Image -OutFile "$ProfilePath\TempImage-$Name.png" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download image. Make sure the URL is correct and accessible."
        Break
    }

    # Save icon from image and cleanup
    try {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        $LoadedImage = [Drawing.Image]::FromFile("$ProfilePath\TempImage-$Name.png")
        $IntPtr = New-Object IntPtr
        $Thumbnail = $LoadedImage.GetThumbnailImage(72, 72, $null, $IntPtr)
        $Bitmap = New-Object Drawing.Bitmap $Thumbnail
        $Bitmap.SetResolution(72, 72);
        $Icon = [System.Drawing.Icon]::FromHandle($Bitmap.GetHicon());
        $File = [IO.File]::Create("$ProfilePath\Edge Profile.ico")
        $Icon.Save($File)
        $File.Close()
        $Icon.Dispose()
        $Bitmap.Dispose()
        $LoadedImage.Dispose()
        $Thumbnail.Dispose()
        Remove-Item "$ProfilePath\TempImage-$Name.png" -Force
    }
    catch {
        Write-Error "Failed to save icon from image. Make sure the image is a valid."
        Break
    }
}

if ($StartMenuShortcut) {
    # Create Start Menu shortcut
    try {
        $TargetPath = $EdgePath
        $ShortcutFile = "C:\ProgramData\Microsoft\Windows\Start Menu\$Name - Edge.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.IconLocation = "$ProfilePath\Edge Profile.ico, 0"
        $Shortcut.Arguments = "--profile-directory=`"$ProfileFolder`""
        $Shortcut.TargetPath = $TargetPath
        $Shortcut.Save()
    }
    catch {
        Write-Error "Failed to create Start Menu shortcut."
        Break
    }
}

Write-Output "Successfully created Edge profile `"$Name`"."