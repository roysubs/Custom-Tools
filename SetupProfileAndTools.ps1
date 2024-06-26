########################################
#
# SetupProfileAndTools.ps1
#
# Author: roysubs@hotmail.com
#
# 2022-05-01 Initial setup
#
########################################
#
# Install Profile Extensions & Custom-Tools.psm1 if available locally, or get them from github.
#
# To load directly from the internet:
#    . iex ((New-Object System.Net.WebClient).DownloadString('tinyurl.com/SetupProfileAndTools')); !!! change this link for this new file !!!
#
# or
#
# To load from a local clone of the project:
#    git clone https://github.com/roysubs/Custom-Tools
#    . ./SetupProfileAndTools.ps1
#
# Create short links with the https://git.io url shortener (no, they no longer accept, so just use tinyurl.com)
#
# Sometimes, Windows Defender updates can block many PowerShell scripts. Here are notes and workarounds:
# https://theitbros.com/managing-windows-defender-using-powershell/
# https://technoresult.com/how-to-disable-windows-defender-using-powershell-command-line/
# https://evotec.xyz/import-module-this-script-contains-malicious-content-and-has-been-blocked-by-your-antivirus-software/
# https://superuser.com/questions/1503345/i-disabled-real-time-monitoring-of-windows-defender-but-a-powershell-script-is
# Even the Chocolatey installer can be affected if Windows Defender is not turned off: https://github.com/chocolatey/choco/issues/2132
# In this case, temporarily disable Windows Defernder (use of 'iex' might be why Defernder throws a tantrum):
#    sc.exe stop|start|query WinDefend   # Note use of sc.exe and not sc (Set-Content alias in PowerShell)
#    Set-MpPreference -DisableRealtimeMonitoring $true (to disable) or $false (to enable)   # RealtimeMonitoring
#
# GitHub access uses TLS. Sometimes (not always) have to specify this:
#   [Net.ServicePointManager]::SecurityProtocol   # To view current settings
#   [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  # Set type to TLS 1.2
#      # Note that the above Tls12 is incompatible with Windows 7 which only has SSL3 and TLS as options.
#   [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls    # Set type to TLS
#
# GitHub pages are cached at the endpoint so after uploading to GitHub, the new data will not be immediately
# available and can take from ~10s to 180s to update. Clearing the DNS Client Cache might help:
#   On PowerShell v2, use: [System.Net.ServicePointManager]::DnsRefreshTimeout = 0;
#   On newer versions of PowerShell, can use: Clear-DnsClientCached
# https://stackoverflow.com/questions/18556456/invoke-webrequest-working-after-server-is-unreachable
#
########################################

function Write-Header {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host 'Auto-Configure Custom Tools (ProfileExtensions.ps1 & Custom-Tools.psm1)' -ForegroundColor Green
    Write-Host 'Will adjust the profile to call ProfileExtensions.ps1 and setup the Modules' -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
}

# Variables, create HomeFix in case of network shares (as always want to use C:\ drive, so get the name (Leaf) from $HOME)
# Need paths that fix the "Edwin issue" i.e. UsernName has changed from the path that in $env:USERPROFILE
# And also to fix the issue with VPN network paths, so  check for "\\" in the profile,    # $UserNameFriendly = $env:UserName
$HomeFix = $HOME
$HomeLeaf = split-path $HOME -leaf   # Just get the correct username in spite of any changes to username!
$WinId = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name   # This returns Hostname\WindowsIdentity, where WindowsIdentity can be different from UserProfile folder name
if ($HomeFix -like "\\*") { $HomeFix = "C:\Users\$(Split-Path $HOME -Leaf)" }
$UserModulePath = "$HomeFix\Documents\WindowsPowerShell\Modules"   # $UserModulePath = "C:\Users\$HomeLeaf\Documents\WindowsPowerShell\Modules"
$UserScriptsPath = "$HomeFix\Documents\WindowsPowerShell\Scripts"
$AdminModulesPath = "C:\Program Files\WindowsPowerShell\Modules"
# The default Modules and Scripts paths are not created by default in Windows
if (!(Test-Path $UserModulePath)) { md $UserModulePath -Force -EA silent | Out-Null }
if (!(Test-Path $UserScriptsPath)) { md $UserScriptsPath -Force -EA silent | Out-Null }

$OSver = (Get-CimInstance win32_operatingsystem).Name
$PSver = $PSVersionTable.PSVersion.Major

function Test-Administrator {
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Add-ToPath {
    param (
        [string]$PathToAdd,
        [Parameter(Mandatory=$true)][ValidateSet('System','User')]      [string]$UserType,
        [Parameter(Mandatory=$true)][ValidateSet('Path','PSModulePath')][string]$PathType
    )
    # Add-ToPath "C:\XXX" 'System' "PSModulePath"
    if ($UserType -eq "User"   ) { $RegPropertyLocation = 'HKCU:\Environment' } # also note: Registry::HKEY_LOCAL_MACHINE\ format
    if ($UserType -eq "System" ) { $RegPropertyLocation = 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' }
    "`nAdd '$PathToAdd' (if not already present) into the $UserType `$$PathType"
    "The '$UserType' environment variables are held in the registry at '$RegPropertyLocation'"
    try { $PathOld = (Get-ItemProperty -Path $RegPropertyLocation -Name $PathType -EA silent).$PathType } catch { "ItemProperty is missing" }
    "`n$UserType `$$PathType Before:`n$PathOld`n"
    $PathOld = $PathOld -replace "^;", "" -replace ";$", ""   # After displaying actual old path, remove leading/trailing ";" (also .trimstart / .trimend)
    $PathArray = $PathOld -split ";" -replace "\\+$", ""      # Create the array, removing network locations???
    if ($PathArray -notcontains $PathToAdd) {
        "$UserType $PathType Now:"   # ; sleep -Milliseconds 100   # Might need pause to prevent text being after Path output(!)
        $PathNew = "$PathOld;$PathToAdd"
        Set-ItemProperty -Path $RegPropertyLocation -Name $PathType -Value $PathNew
        Get-ItemProperty -Path $RegPropertyLocation -Name $PathType | select -ExpandProperty $PathType
        if ($PathType -eq "Path") { $env:Path += ";$PathToAdd" }                  # Add to Path also for this current session
        if ($PathType -eq "PSModulePath") { $env:PSModulePath += ";$PathToAdd" }  # Add to PSModulePath also for this current session
        "`n$PathToAdd has been added to the $UserType $PathType"
    }
    else {
        "'$PathToAdd' is already in the $UserType $PathType. Nothing to do."
    }
}

# Update PSModulePath (User), check/add the default Modules location $UserModulePath to the user PSModulePath
Add-ToPath $UserModulePath "User" "PSModulePath" | Out-Null   # Add the correct User Modules path to PSModulePath

# Update Path (User), check/add the default Scripts location $UserScriptsPath to the user Path
Add-ToPath $UserScriptsPath "User" "Path" | Out-Null    # Add the correct User Scripts path to Path

function ThrowScriptErrorAndStop {
    ""
    throw "This is not an error. Using the 'throw' command here to halt script execution`nas 'return' / 'exit' have issues when run with Invoke-Expression from a URL ..."
}

function Confirm-Choice {
    param ( [string]$Message )
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes";
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No";
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no);
    $caption = ""   # Did not need this before, but now getting odd errors without it.
    $answer = $host.ui.PromptForChoice($caption, $message, $choices, 0)   # Set to 0 to default to "yes" and 1 to default to "no"
    switch ($answer) {
        0 {return 'yes'; break}  # 0 is position 0, i.e. "yes"
        1 {return 'no'; break}   # 1 is position 1, i.e. "no"
    }
}

$unattended = $false   # default condition is to ask user for input
$confirm = 'Do you want to continue?'   # Apart from unattended question, this is used for all other $Message values in Confirm-Choice.

if ((Get-ExecutionPolicy -Scope LocalMachine) -eq "Restricted") {
    "Get-ExecutionPolicy -Scope LocalMachine => Restricted"
    "To allow scripts to run, you need to change the Execution Policy from an Administrator console."
    "'Set-ExecutionPolicy RemoteSigned' or 'Set-ExecutionPolicy Unrestricted' will permit this."
    "One of these is required to run `$profile and additional scripts in this configuration."
    ""
    #!!!!! if ($(Confirm-Choice "Stop this configuration run until the ExecutionPolicy is configured?`nSelect 'n' to continue anyway (expect errors).") -eq "no") { $unattended = $true }
}

function Write-Wrap {
    <#
    .SYNOPSIS
    Wraps a string or an array of strings at the console width so that no word is broken at line enddings and neatly folds to multiple lines
    https://stackoverflow.com/questions/1059663/is-there-a-way-to-wordwrap-results-of-a-powershell-cmdlet#1059686
    .PARAMETER chunk
    A string or an array of strings
    .EXAMPLE
    Write-Wrap -chunk $string
    .EXAMPLE
    $string | Write-Wrap
    #>
    [CmdletBinding()]Param(
        [parameter(Mandatory=1, ValueFromPipeline=1, ValueFromPipelineByPropertyName=1)] [Object[]]$chunk
    )
    PROCESS {
        $Lines = @()
        foreach ($line in $chunk) {
            $str = ''
            $counter = 0
            $line -split '\s+' | % {
                $counter += $_.Length + 1
                if ($counter -gt $Host.UI.RawUI.BufferSize.Width) {
                    $Lines += ,$str.trim()
                    $str = ''
                    $counter = $_.Length + 1
                }
                $str = "$str$_ "
            }
            $Lines += ,$str.trim()
        }
        $Lines
    }
}

####################
#
# Note on using Write-Host: Write-Host does not play nicely with the pipeline; it ignores the pipeline and just fires
# things onto the screen. The pipeline does things in a different fashion waiting for the pipeline to close etc. So
# Write-Host means that you will see some outputs muddled up. i.e. Outputs from the pipeline can happen after a
# Write-Host that comes later on in a script. There are good workarounds (mainly to make a function like Write-Host
# that is pipeline compliant), and they might be worth using, but for now I'm just sticking with Write-Host and using
# the "pipeline cludge" which is to put " | Out-Host" on the end of a few pipeline commands that might get jumbled up
# in ordering to make them adhere to Out-Host sequential ordering. i.e. With " | Out-Host" on the end of a Pipeline
# Cmdlet line, then the output will be ordered sequentially along with Write-Host commands.
# Alternatives (pipeline compliant colour outputting Cmdlets):
# https://stackoverflow.com/questions/59220186/usage-of-write-output-is-very-unreliable-compared-to-write-host/59228534#59228534
# https://stackoverflow.com/questions/2688547/multiple-foreground-colors-in-powershell-in-one-command/46046113#46046113
# https://jdhitsolutions.com/blog/powershell/3462/friday-fun-out-conditionalcolor/
# https://www.powershellgallery.com/packages?q=write-coloredoutput
#
# The best approach might be to bypass Write-Host with a custom function (functions take priority over Cmdlets as
# shown here): https://stackoverflow.com/questions/33747257/can-i-override-a-powershell-native-cmdlet-but-call-it-from-my-override
#
####################

""
# Test the path that this script is running from; if this path is $null, then it must have started by an Invoke-Expression from a URL.
$SetupPath = $MyInvocation.MyCommand.Path   # echo $SetupPath
# $UrlConfig            = 'https://raw.githubusercontent.com/roysubs/Custom-Tools/main/SetupProfileAndTools.ps1'   # Check master/main confusion in case project changes
$UrlConfig            = 'https://raw.githubusercontent.com/roysubs/Custom-Tools/main/SetupProfileAndTools.ps1'   # Check master/main confusion in case project changes
$UrlProfileExtensions = 'https://raw.githubusercontent.com/roysubs/Custom-Tools/main/ProfileExtensions.ps1'
$UrlCustomTools       = 'https://raw.githubusercontent.com/roysubs/Custom-Tools/main/Custom-Tools.psm1'
$UrlCustomExternal    = 'https://raw.githubusercontent.com/roysubs/Custom-Tools/main/Custom-External.psm1'

if ($SetupPath -eq $null) {
    "Scripts are being downloaded and installed from internet (github)."
    # This is case when $SetupPath is null, i.e. the script was run via the web, so have to download all scripts.
    # (New-Object System.Net.WebClient).DownloadString('https://bit.ly/2R7znLX') | Out-File "$env:TEMP\SetupProfileAndTools.ps1"
    # if (Test-Path "$ScriptSetupPath\SetupProfileAndTools.ps1") { rm "$env:TEMP\SetupProfileAndTools.ps1" -Force }
    # if (Test-Path "$ScriptSetupPath\ProfileExtensions.ps1")    { rm "$env:TEMP\ProfileExtensions.ps1" -Force }
    # if (Test-Path "$ScriptSetupPath\Custom-Tools.psm1")        { rm "$env:TEMP\Custom-Tools.psm1" -Force }
    # if (Test-Path "$ScriptSetupPath\Custom-External.psm1")     { rm "$env:TEMP\Custom-External.psm1" -Force }
    rm "$env:TEMP\SetupProfileAndTools.ps1" -Force -EA Silent
    rm "$env:TEMP\ProfileExtensions.ps1" -Force -EA Silent
    rm "$env:TEMP\Custom-Tools.psm1" -Force -EA Silent
    rm "$env:TEMP\Custom-External.psm1" -Force -EA Silent
    (New-Object System.Net.WebClient).DownloadString($UrlConfig)            | Out-File "$env:TEMP\SetupProfileAndTools.ps1" -Force
    (New-Object System.Net.WebClient).DownloadString($UrlProfileExtensions) | Out-File "$env:TEMP\ProfileExtensions.ps1" -Force
    (New-Object System.Net.WebClient).DownloadString($UrlCustomTools)       | Out-File "$env:TEMP\Custom-Tools.psm1" -Force
    (New-Object System.Net.WebClient).DownloadString($UrlCustomExternal)    | Out-File "$env:TEMP\Custom-External.psm1" -Force
    $SetupPath = "$env:TEMP\SetupProfileAndTools.ps1"
    $ScriptSetupPath = Split-Path $SetupPath   # Copy the files to TEMP as staging area for local or downloaded files
}
else {
    "Scripts are being installed locally from:   $SetupPath"
    # If the path is not null, then the script was run from the filesystem, so the scripts should be here.
    # First, test if all scripts are here, if they are not, then no point in continuing.
    $ScriptSetupPath = Split-Path $SetupPath   # Copy the files to TEMP as staging area for local or downloaded files
    if ((Test-Path "$ScriptSetupPath\SetupProfileAndTools.ps1") `
     -and (Test-Path "$ScriptSetupPath\ProfileExtensions.ps1") `
     -and (Test-Path "$ScriptSetupPath\Custom-Tools.psm1") `
     -and (Test-Path "$ScriptSetupPath\Custom-External.psm1") `
     -and ($ScriptSetupPath -ne $env:TEMP)) {
            # If the running scripts are not in TEMP, then copy them there and overwrite
            # Slight issue! When elevating to admin, the called scripts are in TEMP(!), so skip copying as will be to same location!
            if (Test-Path "$ScriptSetupPath\SetupProfileAndTools.ps1") { Copy-Item "$ScriptSetupPath\SetupProfileAndTools.ps1" "$env:TEMP\SetupProfileAndTools.ps1" -Force }
            if (Test-Path "$ScriptSetupPath\ProfileExtensions.ps1")    { Copy-Item "$ScriptSetupPath\ProfileExtensions.ps1"    "$env:TEMP\ProfileExtensions.ps1" -Force }
            if (Test-Path "$ScriptSetupPath\Custom-Tools.psm1")        { Copy-Item "$ScriptSetupPath\Custom-Tools.psm1"        "$env:TEMP\Custom-Tools.psm1" -Force }
            if (Test-Path "$ScriptSetupPath\Custom-External.psm1")     { Copy-Item "$ScriptSetupPath\Custom-External.psm1"     "$env:TEMP\Custom-External.psm1" -Force }
    }
    else {
        "Scripts not found."
        "   - SetupProfileAndTools.ps1"
        "   - ProfileExtensions.ps1"
        "   - Custom-Tools.psm1"
        "   - Custom-External.psm1"
        "Make sure that all required scripts are available."
        "Exiting script..."
    }
}

####################
#
# Self-elevate script if required.
#
####################

####################
#
# Use the users Module folder (i.e. do not use C:\ProgramData\WindowsPowerShell\Modules)
# as it is important to keep everything in user space instead of requiring elevation, but
# keep these notes here for reference.
#
# \\ad. ing. net\WPS\NL\P\UD\200024\YA6 4UG\Home\My Documents\WindowsPowerShell\Modules
# Redirect to => C:\Users\YA6 4UG\Documents\WindowsPowerShell\Modules
# Profile loading times are slow when running on an ING EndPoint connection.
# Laptop without EndPoint => PowerShell loads in <1 sec
# Laptop with EndPoint    => PowerShell loads in 6.5 to 9 sec
# Laptop with EndPoint    => PowerShell loads in 3.5 sec (if move rofile extensions to C:\ locally)
# Laptop with EndPoint    => PowerShell loads in 1 sec (if also load C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1)
# Plan is then:
#    a) Try to move to profile.ps1 under System32 (will only work if TestAdministrator -eq True)
#    b) Use default user folder if cannot do that. i.e. \\ad.ing.net\WPS\NL\P\UD\200024\YA64UG\Home\My Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
# But this is not a problem because b) is only a problem if on a laptop with VPN because if on a server etc, the default user folder is fine!
# So, if on my laptop:
#    Profile = [C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1] + [C:\User\YA64UG\WindowsPowerShell\profile.ps1]
#        if (Test-File $Profile) { rename $Profile $Profile-disabled }
#    Modules = [C:\Users\YA64UG\Documents\WindowsPowerShell\Modules]
#        Check that this is in $env:PSModulePath
#        Delete the H: version of Modules completely!
#
# 1. if (Test-Administrator -eq True) => ask if want Admin install C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1 or Default User install
#        Suggest to only do an Admin install on personal laptop
# 2.
# 3. Scripts => C:\Users\YA64UG\Documents\WindowsPowerShell\Scripts
#        Make sure that this is on the path
# mklinks => C:\PS\ (so profile will be at C:\PS\profile.ps1), C:\PS\Scripts, C:\PS\Modules
# No, can't do this as might not be admin!
# Just leave at normal locations but have go definitions to jump to these locations.
#
# # To get around the ING problem with working from home.
# By mirroring the normal profile location, then everything else will work normally.
# Profile extensions will be placed in the local profile location rather than the ING one.
# "Microsoft.PowerShell_profile.ps1_extensions.ps1"
# By only appyling this to systems from ING, everything else will work normally.
# So the below is what should be pushed to C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1
#
# # # # if ($env:USERNAME -eq "ya64ug") {
# if ($(get-ciminstance -class Win32_ComputerSystem).PrimaryOwnerName -eq "ING")
#     $PROFILE.CurrentUserAllHosts = "$HomeFix\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"
#     $PROFILE.CurrentUserCurrentHost = "$HomeFix\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"
#     . $PROFILE
#     cd $HomeFix
# }
#
# $PROFILE = "$HomeFix\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"
#
# Note that the above breaks the $PROFILE definition!
# ($PROFILE).GetType()
# IsPublic IsSerial Name                                     BaseType
# -------- -------- ----                                     --------
# True     True     String                                   System.Object
#
# $PROFILE | Format-List * -Force
# AllUsersAllHosts       : C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1
# AllUsersCurrentHost    : C:\Windows\System32\WindowsPowerShell\v1.0\Microsoft.PowerShell_profile.ps1
# CurrentUserAllHosts    : \\ad.ing.net\WPS\NL\P\UD\200024\YA64UG\Home\My Documents\WindowsPowerShell\profile.ps1
# CurrentUserCurrentHost : \\ad.ing.net\WPS\NL\P\UD\200024\YA64UG\Home\My Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
# Length                 : 107
#
# But, after the change above, making $PROFILE a single string:
# $PROFILE | Format-List * -Force
# C:\Users\ya64ug\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
#
# So instead, just change those parts:
# $PROFILE.CurrentUserAllHosts = "$HomeFix\Documents\WindowsPowerShell\profile.ps1"
# $PROFILE.CurrentUserCurrentHost = "$HomeFix\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"
#
# More here: https://devblogs.microsoft.com/scripting/understanding-the-six-powershell-profiles/
#
####################

Write-Header

# Elevation will restart the script so don't ask this question until after that point
#!!!!! if ($(Confirm-Choice "Prompt all main action steps during setup?`nSelect 'n' to make all actions unattended.") -eq "no") { $unattended = $true }
$unattended = $true



####################
#
# Main Script
#
####################



$start_time = Get-Date   # Put following lines at end of script to time it
# $hr = (Get-Date).Subtract($start_time).Hours ; $min = (Get-Date).Subtract($start_time).Minutes ; $sec = (Get-Date).Subtract($start_time).Seconds
# if ($hr -ne 0) { $times += "$hr hr " } ; if ($min -ne 0) { $times += "$min min " } ; $times += "$sec sec"
# "Script took $times to complete"   # $((Get-Date).Subtract($start_time).TotalSeconds)

# Set-ExecutionPolicy RemoteSigned
Write-Host ""
Write-Host "Try to set the execution policy to RemoteSigned for this system" -ForegroundColor Yellow -BackgroundColor Black
try {
    Set-ExecutionPolicy RemoteSigned -Force   # Need
} catch {
    Write-Host "`nWARNING: 'Set-ExecutionPolicy RemoteSigned' failed to execute." -ForegroundColor Yellow
    Write-Host "This is often due to Group Policy restrictions on corporate builds.`n" -ForegroundColor Yellow
    Write-Host "'Get-ExecutionPolicy -List' (show current execution policy list):`n"
    Get-ExecutionPolicy -List | ft
}

# Set TLS for GitHub compatibility
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls } catch { }     # Windows 7 compatible
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }   # Windows 10 compatible

# SetupProfileAndTools.ps1 is much simpler than the old setup. Just install the Profile Extensions and Modules
# Things like chocolatey / boxstarter setup will be left out.

Write-Host ""
Write-Host 'The object of this script is to configure basic System and PowerShell configurations.'
Write-Host 'Avoiding extensive modifications and sticking to core functionality.'
Write-Host ""
Write-Host "# The current script is used to chain together these system configurations" -ForegroundColor DarkGreen
Write-Host "iwr '$UrlConfig' | select -expandproperty content | more" -ForegroundColor Magenta
Write-Host "(New-Object System.Net.WebClient).DownloadString('$UrlConfig')"
Write-Host "# The Profile Extension script extends `$profile" -ForegroundColor DarkGreen
Write-Host "iwr '$UrlProfileExtensions' | select -expandproperty content | more" -ForegroundColor Magenta
Write-Host "(New-Object System.Net.WebClient).DownloadString('$UrlProfileExtensions')"
Write-Host "# Custom-Tools.psm1 Module to make useful generic functions available to console." -ForegroundColor DarkGreen
Write-Host "iwr '$UrlCustomTools' | select -expandproperty content | more" -ForegroundColor Magenta
Write-Host "(New-Object System.Net.WebClient).DownloadString('$UrlCustomTools')"
Write-Host "# Custom-External.psm1 Module to make useful external functions from the web available to console." -ForegroundColor DarkGreen
Write-Host "iwr '$UrlCustomExternal' | select -expandproperty content | more" -ForegroundColor Magenta
Write-Host "(New-Object System.Net.WebClient).DownloadString('$UrlCustomExternal')"
Write-Host ""
Write-Host 'Please review the above URLs to check that the changes are what you expect before continuing.'
Write-Host ""

#!!!!! if ($unattended -eq $false) { if ($(Confirm-Choice $confirm) -eq "no") { ThrowScriptErrorAndStop } }



####################
#
# Setup Profile Extensions and Custom Tools Module
#
####################

Write-Host ''
Write-Host ''
Write-Host "`n========================================" -ForegroundColor Green
Write-Host ''
Write-Host 'Configure Custom-Tools and setup Profile Extensions in $Profile.' -ForegroundColor Green
Write-Host ''
Write-Host 'Check for the profile extensions in same folder as $Profile. If not present, the' -ForegroundColor Yellow
Write-Host 'latest profile extensions will be downloaded from Gist. A single line will then be' -ForegroundColor Yellow
Write-Host 'added into $Profile to load the profile extensions from $Profile into all new consoles.' -ForegroundColor Yellow
Write-Host ''
Write-Host 'To force the latest profile extensions, either rerun this script which will overwrite it' -ForegroundColor Yellow
Write-Host 'or delete the profile extensions and it will be downloaded on opening the next console.' -ForegroundColor Yellow
Write-Host ''
Write-Host 'The reason for having the profile extensions separately is to keep the profile clean and' -ForegroundColor Yellow
Write-Host 'to keep additions synced against a known working online copy.' -ForegroundColor Yellow
if ($OSver -like "*Server*") {
    Write-Host ''
    Write-Host "The Operating System is of type 'Server' so profile extensions will *not* load" -ForegroundColour Yello -BackgroundColor Black
    Write-Host "by default in any console. Profile extensions can be loaded on demand by running" -ForegroundColour Yellow -BackgroundColor Black
    Write-Host "the 'Enable-Extensions' function that will now be setup in `$profile." -ForegroundColour Yellow -BackgroundColor Black
}
Write-Host ''



Write-Host "========================================`n" -ForegroundColor Green
Write-Host ''
Write-Host ''
Write-Host "Update and reinstall Custom-Tools.psm1 Module." -ForegroundColor Yellow -BackgroundColor Black
Write-Host "The Module contains many generic functions (including some required later in this setup)."
Write-Host "View functions in the module with 'myfunctions' / 'mods' / 'mod custom-tools' or view module contents with:"
Write-Host "   get-command -module custom-tools" -ForegroundColor Yellow

$CustomTools = "$HomeFix\Documents\WindowsPowerShell\Modules\Custom-Tools\Custom-Tools.psm1"
if ([bool](Get-Module Custom-Tools -ListAvailable) -eq $true) {
    if ($PSver -gt 4) { Uninstall-Module Custom-Tools -EA Silent -Force -Verbose }   # Uninstall-Module is only in PS v4+
    else { "Need to be running PS v5 or higher to run Uninstall-Module" }
}
if (!(Test-Path (Split-Path $CustomTools))) { New-Item (Split-Path $CustomTools) -ItemType Directory -Force }   # Create folder if required
if (Test-Path ($CustomTools)) { rm $CustomTools }   # Delete old version of the .psm1 if it was already there
"$SetupPath is currently running script."
"$(Split-Path $SetupPath)\Custom-Tools.psm1 will be used to load Custom-Tools Module."

if ($SetupPath -eq $null) {
    # Case when scripts were downloaded
    # try { (New-Object System.Net.WebClient).DownloadString("$UrlCustomTools") | Out-File $CustomTools ; echo "Downloaded Custom-Tools.psm1 Module from internet ..." }
    # catch { Write-Host "Failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
    Copy-Item "$ScriptSetupPath\Custom-Tools.psm1" $CustomTools -Force
}
else {
    # Try to use the version in same folder as SetupProfileAndTools
    Copy-Item "$ScriptSetupPath\Custom-Tools.psm1" $CustomTools -Force
    Write-Host "$ScriptSetupPath\Custom-Tools.psm1 was copied to Custom-Tools Path: $CustomTools"
    Write-Host "Using local version of Custom-Tools.psm1 from $SetupPath ..."
}

Import-Module Custom-Tools -Force -Verbose    # Don't require full path, it should search for it in standard PSModulePaths
$x = ""; foreach ($i in (get-command -module Custom-Tools).Name) {$x = "$x, $i"} ; "" ; Write-Wrap $x.trimstart(", ") ; ""

Write-Host ''
Write-Host ''
Write-Host "========================================`n" -ForegroundColor Green
Write-Host ''
Write-Host ''
Write-Host "Update and reinstall Custom-External.psm1 Module." -ForegroundColor Yellow -BackgroundColor Black
Write-Host "The Module dynamically links to many external scripts and functions."
Write-Host "View functions in the module with 'myfunctions' / 'mods' / 'mod custom-external' or view module contents with:"
Write-Host "   get-command -module custom-external" -ForegroundColor Yellow

$CustomExternal = "$HomeFix\Documents\WindowsPowerShell\Modules\Custom-External\Custom-External.psm1"
if ([bool](Get-Module Custom-External -ListAvailable) -eq $true) {
    if ($PSver -gt 4) { Uninstall-Module Custom-External -EA Silent -Force -Verbose }   # Uninstall-Module is only in PS v4+
    else { "Need to be running PS v5 or higher to run Uninstall-Module" }
}
if (!(Test-Path (Split-Path $CustomExternal))) { New-Item (Split-Path $CustomExternal) -ItemType Directory -Force }   # Create folder if required
if (Test-Path ($CustomExternal)) { rm $CustomExternal }   # Delete old version of the .psm1 if it was already there
"$SetupPath is currently running script."
"$(Split-Path $SetupPath)\Custom-Tools.psm1 will be used to load Custom-Tools Module."

if ($SetupPath -eq $null) {
    # Case when scripts were downloaded
    # try { (New-Object System.Net.WebClient).DownloadString("$UrlCustomTools") | Out-File $CustomTools ; echo "Downloaded Custom-Tools.psm1 Module from internet ..." }
    # catch { Write-Host "Failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
    Copy-Item "$ScriptSetupPath\Custom-External.psm1" $CustomExternal -Force
}
else {
    # Try to use the version in same folder as SetupProfileAndTools
    Copy-Item "$ScriptSetupPath\Custom-External.psm1" $CustomExternal -Force
    Write-Host "$ScriptSetupPath\Custom-External.psm1 was copied to Custom-Tools Path: $CustomExternal"
    Write-Host "Using local version of Custom-External.psm1 from $SetupPath ..."
}

Import-Module Custom-External -Force -Verbose    # Don't require full path, it should search for it in standard PSModulePaths
$x = ""; foreach ($i in (get-command -module Custom-External).Name) {$x = "$x, $i"} ; "" ; Write-Wrap $x.trimstart(", ") ; ""




#!!!!! if ($unattended -eq $false) { if ($(Confirm-Choice $confirm) -eq "no") { ThrowScriptErrorAndStop } }




$ProfileFolder     = $Profile | Split-Path -Parent
$ProfileFile       = $Profile | Split-Path -Leaf
$ProfileExtensions = Join-Path $ProfileFolder "$($ProfileFile)_extensions.ps1"
Write-Host "Profile           : $Profile"
Write-Host "ProfileExtensions : $ProfileExtensions"
Write-Host ""
Write-Host "Note that the Profile path below is determined by the currently running 'Host'."
Write-Host "i.e. The host is usually either the PowerShell Console, or PowerShell ISE, or"
Write-Host "Visual Studio Code, each of which will have their own Profile path which you can"
Write-Host "see by typing `$Profile from within that given host. Other hosts can exist, such"
Write-Host "as PowerShell Core running under Linux, but the above are the most usual on Wwindows."
Write-Host ""
Write-Host "You can see more information on the current host by typing '`$host' at the console:"
$host   # Will show the current host

# Create $Profile folder and file if they do not exist
if (!(Test-Path $ProfileFolder)) { New-Item -Type Directory $ProfileFolder -Force }
if (!(Test-Path $Profile)) { New-Item -Type File $Profile -Force }

# Create a backup of the extensions in case user has modified this directly.
if (Test-Path ($ProfileExtensions)) {
    Write-Host "`nCreating backup of existing profile extensions in case download fails ..."
    Move-Item -Path "$($ProfileExtensions)" -Destination "$($ProfileExtensions)_$(Get-Date -format "yyyy-MM-dd__HH-mm-ss").txt"
}

Write-Host "`nGet latest profile extensions (locally if available or download) ..."

if ($SetupPath -eq $null) {   # If SetupPath is null, files should still be in env:TEMP from previous download
    # try { (New-Object System.Net.WebClient).DownloadString("$UrlProfileExtensions") | Out-File $ProfileExtensions -Force ; echo "Downloaded profile extensions from internet ..." }
    # catch { Write-Host "Failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
    Copy-Item "$ScriptSetupPath\ProfileExtensions.ps1" $ProfileExtensions -Force
} else {
    # First try to use the version in same folder as SetupProfileAndTools
    Copy-Item "$ScriptSetupPath\ProfileExtensions.ps1" $ProfileExtensions -Force
    Write-Host "$ScriptSetupPath\ProfileExtensions.ps1 was copied to $ProfileExtensions"
    Write-Host "Using local version of Profile Extensions from $SetupPath ..."
}
# # If still have nothing, try to download from Github
# else {
#     try { (New-Object System.Net.WebClient).DownloadString("$UrlProfileExtensions") | Out-File $UrlProfileExtensions }
#     catch { Write-Host "Failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
# }

# If the script was run from a url, then download latest copies from there.
# For working offline, use the latest profile extensions locally if available by copying to TEMP and using same logic.

# if ($null -eq $SetupPath ) {
#     if (Test-Path "$($env:TEMP)\ProfileExtensions.ps1") {
#         Copy-Item "$($env:TEMP)\ProfileExtensions.ps1" $ProfileExtensions -Force
#         echo "`nUsing local profile extensions from $($env:TEMP)\ProfileExtensions.ps1..."
#     } else {
#         # If $ProfileExtensions was not there, just get from online
#         try { (New-Object System.Net.WebClient).DownloadString("$UrlProfileExtensions") | Out-File $ProfileExtensions ; echo "Downloaded profile extensions from internet ..."}
#         catch { Write-Host "Failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
#     }
# }

if (Test-Path ($ProfileExtensions)) {
    Write-Host "`nDotsourcing new profile extensions into this current session ..."
    . $ProfileExtensions
}
else {
    Write-Host "`nProfile extensions could not be loaded locally (missing) or downloaded from internet."
    Write-Host "If downloaded from internet, check internet/TLS settings before retrying."
}
# If running on a server OS, or "Administrator" or hostname is "SJ*CAL", do not autoload extensions.
# In that case, put a function into $profile to enable:   . Enable-Extensions

if ($OSver -like "*Server*") {

    # Remove the Enable-Extensions line by getting the content *minus* the line to remove "-NotMatch" and adding it back into the profile
    Write-Host "`nRemoving profile extensions handler line from `$Profile to ensure that it is positioned at the end of the script ..."
    Set-Content -Path $profile -Value (Get-Content -Path $profile | Select-String -Pattern '^if \(\!\(Test-Path \("\$\(\$Profile\)_extensions\.ps1\"\)\)\) \{ try { \(New' -NotMatch)
    Set-Content -Path $profile -Value (Get-Content -Path $profile | Select-String -Pattern '^function Enable-Extensions { if' -NotMatch)
    # Set-Content -Path $profile -Value (Get-Content -Path $profile | Select-String -Pattern '^if \($MyInvocation.InvocationName -eq "Enable-Extensions"\)' -NotMatch)

    # Append the Enable-Extensions line to end of $profile
    Write-Host "`nAdding updated profile extensions handler line to `$Profile ...`n"

    # Found that System.Net.WebClient seemed to cause a slowdown during profile loading, so replaced this by a message on how to install if profile extensions are missing.
    # This old syntax worked perfectly though and was tricky to work out so keep it here.
    # $ProfileExtensionsHandler  = "if (!(Test-Path (""`$(`$Profile)_extensions.ps1""))) { try { (New-Object System.Net.WebClient).DownloadString('$UrlProfileExtensions') | Out-File ""`$(`$Profile)_extensions.ps1"" } "
    # $ProfileExtensionsHandler += 'catch { "Could not download profile extensions, check internet/TLS before opening a new console." } } ; '
    # $ProfileExtensionsHandler += '. "$($Profile)_extensions.ps1" -EA silent'
    # Add-Content -Path $profile -Value $ProfileExtensionsHandler -PassThru
    # $ProfileExtensionsHandler = 'if ($MyInvocation.InvocationName -eq "Enable-Extensions") { "`nWarning: Must dotsource Enable-Extensions or it will not be added!`n`n. Enable-Extensions`n" } else { . "$($Profile)_extensions.ps1" -EA silent } } }'

    $ProfileExtensionsHandler = "function Enable-Extensions { if (Test-Path (""`$(`$Profile)_extensions.ps1"")) { echo `"``nProfile extensions not found, install as follows (check internet/TLS if this fails):`n`n# (New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/roysubs/Custom-Tools/main/ProfileExtensions.ps1') | Out-File `"`$(`$Profile)_extensions.ps1`"`n`" } "
    Add-Content -Path $profile -Value $ProfileExtensionsHandler -PassThru | Out-Null
    $ProfileExtensionsHandler = "if (Test-Path (`"`$(`$Profile)_extensions.ps1`")) { . `"`$(`$Profile)_extensions.ps1`" }"
    Add-Content -Path $profile -Value $ProfileExtensionsHandler -PassThru | Out-Null
    Write-Host ""
    Write-Host "The profile extensions handler has *not* been added to `$Profile as the Operating System" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "is of type 'Server'. However, the 'Enable-Extensions' function has been added to `$Profile" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "and can be loaded at any time using by dotsourcing into the current session:" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "   . Enable-Extensions" -ForegroundColor Cyan -BackgroundColor Black
    Write-Host ""

} else {

    # Want to put these lines at the end of $profile, so strip the matching lines and rewrite to the same file
    # Set-Content -Path "C:\myfile.txt" -Value (Get-Content -Path "C:\myfile.txt" | Select-String -Pattern 'H\|159' -NotMatch)
    Write-Host "`nRemoving profile extensions handler line from `$Profile to ensure that it is positioned at the end of the script ..."
    # Set-Content -Path $profile -Value (Get-Content -Path $profile | Select-String -Pattern '^\[Net\.ServicePointManager\]::SecurityProtocol' -NotMatch)
    # Remove the Tls setting as not compatible with PowerShell v2

    # Set-Content -Path $profile -Value (Get-Content -Path $profile | Select-String -Pattern '^\$UrlProfileExtensions \= ' -NotMatch) -EA SilentlyContinue
    Set-Content -Path $profile -Value (Get-Content -Path $profile | Select-String -Pattern '^if \(\!\(Test-Path \("\$\(\$Profile\)_extensions\.ps1\"\)\)\) \{ ' -NotMatch)
    Set-Content -Path $profile -Value (Get-Content -Path $profile | Select-String -Pattern '^if \(Test-Path \("\$\(\$Profile\)_extensions\.ps1\"\)\) \{ ' -NotMatch)
    # Then append the lines in the correct order to the end of $profile. Note "-PassTru" switch to display the result as well as writing to file
    Write-Host "`nAdding updated profile extensions handler line to `$Profile ...`n"

    # $ProfileExtensionsHandler = "`$UrlProfileExtensions = ""$UrlProfileExtensions"" ; "
    # $ProfileExtensionsHandler  = "if (!(Test-Path (""`$(`$Profile)_extensions.ps1""))) { echo `"Profile extensions not found, install as follows (check internet/TLS if this fails):``n(New-Object System.Net.WebClient).DownloadString('$UrlProfileExtensions') | Out-File `"`$(`$Profile)_extensions.ps1`"`" }"

    # Found that the System.Net.WebClient was the source of big slowdown at profile loading, so replaced this just by a message on how to install if required.
    # Keep this here as template though as was hard to work out how to embed things properly.
    # $ProfileExtensionsHandler  = "if (!(Test-Path (""`$(`$Profile)_extensions.ps1""))) { try { (New-Object System.Net.WebClient).DownloadString('$UrlProfileExtensions') | Out-File ""`$(`$Profile)_extensions.ps1"" } "
    # $ProfileExtensionsHandler += 'catch { "Could not download profile extensions, check internet/TLS before opening a new console." } } ; '
    # $ProfileExtensionsHandler += '. "$($Profile)_extensions.ps1" -EA silent'

    $ProfileExtensionsHandler = ""   # Reset in case this script has already been dot-sourced
    $ProfileExtensionsHandler = "if (!(Test-Path (`"`$(`$Profile)_extensions.ps1`"))) { echo `"``nProfile extensions not found, install as follows (check internet/TLS if this fails):``n``n# (New-Object System.Net.WebClient).DownloadString('$UrlProfileExtensions') | Out-File ``""```$(```$Profile)_extensions.ps1``""``n`" }"
    Add-Content -Path $profile -Value $ProfileExtensionsHandler -PassThru | Out-Null
    $ProfileExtensionsHandler = 'if (Test-Path ("$($Profile)_extensions.ps1")) { . "$($Profile)_extensions.ps1" }'   # No need for `" `$ etc as using ''
    Add-Content -Path $profile -Value $ProfileExtensionsHandler -PassThru | Out-Null

    Write-Host ""
    Write-Host "The profile extensions handler has been added to `$Profile and so will be" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "loaded by default in all console sessions." -ForegroundColor Yellow -BackgroundColor Black
    Write-Host ""
    # Considerations for building strings: https://powershellexplained.com/2017-11-20-Powershell-StringBuilder/
    # Alternatively, doing three Add-Content lines with -NoNewLine would also work fine.
    # Added ErrorAction (EA) SilentlyContinue to suppress errors if cannot reach the URL
    # Removed the Tls setting as this is not compatible with PowerShell v2
    # Add-Content -Path $profile -Value "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" -PassThru
}


# Write-Host ""
# if ($unattended -eq $false) { if ($(Confirm-Choice $confirm) -eq "no") { ThrowScriptErrorAndStop } }



Write-Host ""
Write-Host ""
if ($unattended -eq $false) {
    Write-Host ""
    Write-Host "Show profile extensions and custom-tools.psm1 release notes?" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host ""
    #!!!!! if ($(Confirm-Choice $confirm) -eq "yes") {
    #!!!!!     if (Get-Command Help-ToolkitConfig -EA Silent) { Help-ToolkitConfig }   # Show the release notes held in Custom-Tools (if that command is loaded/available)
    #!!!!! }
}

Write-Host ""
Write-Host ""
Write-Host "Run 'Help-ToolkitConfig' to review the above notes."
Write-Host ""
Write-Host "Run 'Help-ToolkitCoreApps' to review important system apps to install."
Write-Host ""
# Clean up TEMP folder - no, don't do this, in case the scripts were deliberately run from this location
# if (Test-Path "$env:TEMP\SetupProfileAndTools.ps1") { Remove-Item "$env:TEMP\SetupProfileAndTools.ps1" -Force }
# if (Test-Path "$env:TEMP\ProfileExtensions.ps1") { Remove-Item "$env:TEMP\ProfileExtensions.ps1" -Force }
# if (Test-Path "$env:TEMP\Custom-Tools.psm1")     { Remove-Item "$env:TEMP\Custom-Tools.psm1" -Force }

$hr = (Get-Date).Subtract($start_time).Hours ; $min = (Get-Date).Subtract($start_time).Minutes ; $sec = (Get-Date).Subtract($start_time).Seconds
if ($hr -ne 0) { $times += "$hr hr " } ; if ($min -ne 0) { $times += "$min min " } ; $times += "$sec sec"
"`nScript took $times to complete.`n"   # $((Get-Date).Subtract($start_time).TotalSeconds)

if ($MyInvocation.InvocationName -eq 'SetupProfileAndTools.ps1') { "`nWarning: Toolkit configuration was not dotsourced, so ProfileExtensions will not be active. Either restart a new PowerShell session or rerun as dotsourced:`n`n   . SetupProfileAndTools.ps1`n" }
