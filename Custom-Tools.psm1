#################### 
# 
# Custom-Tools.psm1
# 2019-11-25 Initial setup
# 2023-12-17 Current Version
#
# Module is installed to the Module folder visible to all users (but can only be modified by Administrators):
#    C:\Program Files\WindowsPowerShell\Modules\Custome-Tools
#
# The Module contains only functions to access on demand as required.
# mods           : View all Modules installed in all PSModulePath folders.
# mod <modname>  : View all cmdlets/functions within a given Module.
# modi <modname> : View detailed info (syntax and type) of each cmdlet/function in a given Module.
# def <command>  : View definitions for any command type: Cmdlet/Function/Alias/ExternalCommand with location and syntax.
# 
# To install the Custom-Tools Module, run the BeginSystemConfig.ps1 script remotely:
#    iex ((New-Object System.Net.WebClient).DownloadString('https://bit.ly/2R7znLX'))
# Quick download of BeginSystemConfig.ps1 + ProfileExtensions.ps1 + Custom-Tools.psm1
#   https://gist.github.com/roysubs/1a5eef75a70065f8f2979ccf2703f322
#   shorturl.at/enBN9
#   iex ((New-Object System.Net.WebClient).DownloadString('shorturl.at/enBN9'))
#
####################

####################
#
# Things that could be implemented ...
# Get-LocalWeather https://joeit.wordpress.com/   # Bit of fun, but pull weather info to console
#
# Not built:
# AddTo-SystemPath, RemoveFrom-SystemPath, AddTo-UserPath, RemoveFrom-UserPath (might not want these, build into the generic 'path' function)
#
# ToDo: Enable network access on all systems
# ToDo: Make all networks private (to make compliant with various things like VPN etc)
# ToDo: Enable RDP on all systems
# ToDo: Must add PSReadLine to PowerShell 5.x on Win 7 etc, by default is not installed. Need it for many of the standard console behaviours, Ctrl+R etc
# Excellent function overview: https://medium.com/@forrestpitz/powershell-boilerplate-e91d2b27c904
# PowerShell DevOps: https://itnext.io/writing-maintainable-powershell-503e5b680ed9
# Advanced Functions: https://improvescripting.com/how-to-write-advanced-functions-or-cmdlets-with-powershell-fast/
# PenTesting Examples: http://www.infosecmatter.com/powershell-commands-for-pentesters/
#
# Chocolatey Automatic Packaging for Maintenance: https://chocolatey.org/docs/automatic-packages
# Chocolatey Workshop, Organizational Use: https://github.com/chocolatey/chocolatey-workshop-organizational-use/
# 
# $args & $input & positional parameters etc: https://stackoverflow.com/questions/2157554/how-to-handle-command-line-arguments-in-powershell
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-5.1#input
# Advanced Functions: https://stackoverflow.com/questions/64630362/powershell-pipeline-compatible-select-string-function
#
# PowerShell extension for VS Code: https://github.com/PowerShell/vscode-powershell/issues/new/choose
# If you open the command pallette (F1 or ctrl + P) then search for snippets plenty are there if the extension is loaded.
# Also note this for the source of the snippets: https://github.com/PowerShell/vscode-powershell/blob/master/snippets/PowerShell.json
#
# ex-path processing in VS Code (type ex-path then select an example)
# https://rkeithhill.wordpress.com/2016/02/17/creating-a-powershell-command-that-process-paths-using-visual-studio-code/
# e.g.
# Modify [CmdletBinding()] to [CmdletBinding(SupportsShouldProcess=$true)]
# $paths = @()
# foreach ($aPath in $Path) {
#     # Resolve any relative paths
#     $paths += $psCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($aPath)
# }
# 
# foreach ($aPath in $paths) {
#     if ($pscmdlet.ShouldProcess($aPath, 'Operation')) {
#         # Process each path
#         
#     }
# }
#
####################

####################
#
# Remember that PowerShell is processed sequentially, so cannot put functions below where they are called.
# If top down coding is wanted in PowerShell, can use a 'Main' function:
#
# function Main {
#     MyFunc1
#     MyFunc2
#     <other code ...>
# }
# 
# function MyFunc {
#     "Hello, World!"
# }
# 
# Main
#
#
# Or using a scriptblock, put the body of your script in a scriptblock at the top:
#
# $block = {
#     Echo "Here is where the main body of my script is"
#     Echo "About to call my function"
#     MultiplyByTwo 5
#     Echo "Done calling my function"
# }  
# function MultiplyByTwo([int] $num) { 2 * $num }
# function MultiplyByTen([int] $num) { 10 * $num }
# 
# # Finally, innvoke the script with a single line at the bottom
# & $block
#
####################




####################
#
# Start of main functions
#
####################

function Update-PowerShellStartup {
    # Console startup times can slow over time. Use the following to generate native
    # images for an assembly and its dependencies and install them in the Native Images Cache.
    # https://stackoverflow.com/questions/59341482/powershell-steps-to-fix-slow-startup
    # https://superuser.com/questions/1212442/powershell-slow-starting-on-windows-10
    # powershell -noprofile -ExecutionPolicy Bypass ( Measure-Command { powershell "Write-Host 1" } ).TotalSeconds

    $env:PATH = [Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
    [AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object {
        $path = $_.Location
        if ($path) { 
            $name = Split-Path $path -Leaf
            Write-Host -ForegroundColor Yellow "`r`nRunning ngen.exe on '$name'"
            ngen.exe install $path /nologo
        }
    }

    Write-Host ""
    Write-Host "Option: Adding powershell.exe to the list of Windows Defender exclusions can speed it up considerably (but might be a risk)."
    Write-Host "Option: Create a shortcut to powershell.exe, right-click on it > properties, go to options tab, click on 'use legacy console'. With legacy on things can be faster."
    Write-Host ""
    Write-Host "Location of PowerShell exe can be found with: (Get-Process -Id `$pid).Path   or   (Get-Command PowerShell.exe).Path"
    Write-Host "For PowerShell 5.1, this is:   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    Write-Host "For PowerShell 7.1, this is:   C:\Program Files\PowerShell\7\pwsh.exe"
    Write-Host ""
    Write-Host "To test PowerShell startup times:"
    Write-Host "From DOS:        powershell -noprofile -ExecutionPolicy Bypass ( Measure-Command { powershell 'Write-Host 1' } ).TotalSeconds"
    Write-Host "From PowerShell: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noprofile -ExecutionPolicy Bypass ( Measure-Command { C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe 'Write-Host 1' } ).TotalSeconds"
    Write-Host ""
}

function Get-Size ($dir) {   # Get-Size using robocopy (about 6x faster than native PowerShell method on reasonably sized folders)
    if ($null -eq $dir) { $dir = "." }   # No path provided, use current location
    $rOutput = &robocopy /l /njh /nfl /ndl /njh $dir dummypath /e /bytes
    $bSize = ($rOutput -cmatch 'Bytes :' -split '\s+')[3]
    "{0:N2} MB" -f ($bSize / 1MB)   # Write-Host "$($dir.FullName)`t$bSize"
}
Set-Alias size Get-Size
Set-Alias sz Get-Size

function Get-SizePS ($dir) {   # Native PowerShell find size (much slower than robocopy version)
    if ($null -eq $dir) { $dir = "." }   # No path provided, use current location
    "{0:N2} MB" -f ((Get-ChildItem "$dir" -Recurse -Force -EA silent | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB)
}

# PowerShell allows '/' in a function name, so following can mimic DOS equivalents (for muscle-memory of typing those commands).
# Note also the handy abbreviation/alias sytnax built into PowerShell for Get-ChildItem (dir/ls/gci)
#    e.g.   dir -ad (-Directory) , dir -af (-Files), dir -ah (-Hidden), -ar (-ReadOnly)
# Also, can shorten any parameter as long as no duplicates to that, so can shorten -Attributes to -at
#    e.g.   -at h  (instead of '-Attributes Hidden'),  -at dir  (instead of '-Attributes Directory')
# Also, can use "!" to logical -NOT a flag, e.g.  -at !h   ('attributes not-hidden', show all files and folders that are NOT hidden)
function dir-all ($name) { dir -Force $name }                         # "dir all", shows all files including hidden (keep this here as "-Force" is not the most obvious syntax)
function dir/ah ($name) { dir -Hidden $name }                    # Mimic DOS dir/ad. Show hideen files/folders. Note also:   dir -Force (for Hidden),   OR,   Where Attributes -like '*Hidden*'
function dir/ad ($name) { dir -Directory $name }                 # Mimic DOS dir/ad. Show directories. Note also:   dir -Force (for Hidden),   OR,   Where Attributes -like '*Hidden*'
function dir/a-d ($name) { dir -File $name }                     # Mimic DOS dir/a-d. Show files, i.e. "attributes of 'NOT directories'"". Note also:   dir -Force | Where Attributes -like '*Hidden*' }
function dir/b ($name) { dir $name | select Name | sort Name }   # Mimic DOS dir/b (bare names only). Also sort by Name.
function dir/s ($name) { dir -Force -Recurse $name }             # Mimic DOS dis/s. Dir with subfolders (-Recurse). Add "-Force" to show also Hidden files.
function dir/p ($name) { dir -Force $name | more }               # Mimic DOS dis/p. Dir page by page. Add "-Force" also to show all files.
function dir/os ($name) { dir $name | sort Length,Name }         # Sort by Size (Length), then by Name.

# https://poshoholic.com/2010/11/11/powershell-quick-tip-creating-wide-tables-with-powershell/
# https://stackoverflow.com/questions/1479663/how-do-i-do-dir-s-b-in-powershell
function dir/w ($name) { cmd.exe /c dir /w }   # DOS dir/w, wide format (no proper equivalent with Get-ChildItem)

function dirq ($name) {   # quick and dirty Size + Name, work in progress
    $out = ""
    function Format-FileSize([int64]$size) {
        if ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
        elseif ($size -gt 1GB) {[string]::Format("{0:0.0} GB", $size / 1GB)}
        elseif ($size -gt 1MB) {[string]::Format("{0:0.0} MB", $size / 1MB)}
        elseif ($size -gt 1KB) {[string]::Format("{0:0.0} KB", $size / 1KB)}
        elseif ($size -gt 0) {[string]::Format("{0:0.0} B", $size)}
        else {""}
    }
    foreach ($i in (dir $folder | sort Length).FullName) {
        if (Test-Path -Path $i -PathType Container) {
            $size = "[D]"
            $size_out = "[D]" 
        }
        else {
            $size = (gci $i | select length).Length
            $size_out = Format-FileSize($size)
            $size_total += $size
        }
        $out += "$size_out`t$(split-path $i -leaf)`n"
    }
    $out += "$(Format-FileSize($size_total)) : Total Size"
    $out
}
# $out.TrimEnd("  :  ")   # trime ous whitespace from either size of ":"

function dirwide ($name) {   # quick and dirty wide listing, work in progress
    $out = ""
    function Format-FileSize([int64]$size) {
        if ($size -gt 1TB) {[string]::Format("{0:0.00}TB", $size / 1TB)}
        elseif ($size -gt 1GB) {[string]::Format("{0:0.0}GB", $size / 1GB)}
        elseif ($size -gt 1MB) {[string]::Format("{0:0.0}MB", $size / 1MB)}
        elseif ($size -gt 1KB) {[string]::Format("{0:0.0}kB", $size / 1KB)}
        elseif ($size -gt 0) {[string]::Format("{0:0.0}B", $size)}
        else {""}
    }

    foreach ($i in (dir $folder | sort Length).FullName) {
        if (Test-Path -Path $i -PathType Container) { $size = "[D]" ; $size_out = "[D]" }
        else { $size = (gci $i | select length).Length ; $size_out = Format-FileSize($size) }
        $out += "$i $size_out  :  "
        # $outlength +=
    }
    $out.TrimEnd("  :  ")
}

function dirpaths ($folder, $filter) {   # Remove header information and just show the full paths
    try { gci -r $folder -Filter $filter | select -expand FullName -EA silent }
    catch { "crapped out!" }
}

# Extending Get-ChildItem
# https://jdhitsolutions.com/blog/powershell/9057/using-powershell-your-way/

function killx () { kill -n explorer; explorer }   # Kill explorer and restart it (for times when it doesn't restart immediately)

# Help Functions ...
# ms (MAN SYNTAX), mm (MAN), mp <cmd> <param> (MAN PARAMETER HELP), me (MAN EXAMPLES), mf (MAN FULL)
function Test-Input ($test) { if ($null -eq $args) { "Must specify input." } }   # Need to use exit for this to exit the calling function
# Note: if you don't know the full command, just 'm get' and you get a list of all matching commands, or help *win* etc
function ms ($cmd) { if (Test-Input $args) { break }; Get-Command $cmd -Syntax }   # ms man/help-syntax, actually a Get-Command option   # or (Get-Command $cmd).Definition
Set-Alias syn ms
function mm ($module) { if ($null -eq $module) { "Must specify a module." ; break } ; Get-Command -Module $module }   # mm : man/help for Module contents
function mparam ($cmd, $parameter) { if ($null -eq $cmd) { "Must specify a command." ; break } ; Get-Help $cmd -Parameter $parameter }   # mp : man/help-parameter, get details for a specific parameter
function me ($cmd) { if ($null -eq $cmd) { "Must specify a command." ; break } ; Get-Help $cmd -Examples | more }   # me : man/help-examples, get examples for a specific command
Set-Alias ex me
function mf ($cmd) { if ($null -eq $cmd) { "Must specify a command." ; break } ; Get-Help $cmd -Full | more }   # mf : man/help-full, Full help, the longest output, like detailed, but expand every parameter property (uneccessary)

# Quick access to Environment Variables and currently defined session Variables
function env { gci env: }   # bit of a daft function, but often forget how to list environment variables
function envgui { rundll32 sysdm.cpl,EditEnvironmentVariables }   # Open Environment Variables dialogue
function vars { ls variable:* }   # also a quick function   # Get-Variable | % { "Name : {0}`r`nValue: {1}`r`n" -f $_.Name,$_.Value }
function getvars { Get-Variable | % { "Name : {0}`r`nValue: {1}`r`n" -f $_.Name,$_.Value } }
# Note on resolving Variables (PowerShell) vs Environment Variables (Legacy DOS) ...
# ls variable:*   *or*  variable   on it's own!
# ls env:*   or   env:   or   env
# https://stackoverflow.com/qu



function Get-CommandsByModule ($usertype) {
    $types = @("Alias", "Function", "Filter", "Cmdlet", "ExternalScript", "Application", "Script", "Workflow", "Configuration")
    if ($null -ne $usertype) { $types = @($usertype)}
    foreach ($type in $types) { New-Variable -Name $type -Value 0 }   # Dynamically generated variables

    function Write-Wrap {
        [CmdletBinding()]
        Param ( 
            [parameter (Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [Object[]] $chunk
        )
        PROCESS {
            $Lines = @()
            foreach ($line in $chunk) {
                $str = ''
                $counter = 0
                $line -split '\s+' | %{
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

    foreach ($mod in Get-Module -ListAvailable) {
        "`n`n####################`n#`n# Module: $mod`n#`n####################`n"
        foreach ($type in $types) {
            $out = ""
            $commands = gcm -Module $mod -CommandType $type | sort
            foreach ($i in $commands) {
                $out = "$out, $i"
            }
            $count = ($out.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count   # Could just count $i but this is 

            if ($count -ne 0) {
                $out = $out.trimstart(", ")
                $out = "`n$($type.ToUpper()) objects [ $count ] >>> $out"
                Write-Wrap $out
                # Example of using New-, Set-, Get-Variable for dynamically generated variables
                Set-Variable -Name $type -Value $((Get-Variable -Name $type).Value + $count)
                # https://powershell.org/forums/topic/two-variables-into-on-variable/
                # "$type Total = $total"
                ""
            }
        }
    }
    ""
    "`n`n####################`n#`n# Commands by type installed on this system`n#`n####################`n"
    foreach ($type in $types) { "Total of type '$type' = $((Get-Variable -Name $type).Value)" }
}


function Get-CommandTypes ($type) { Get-Module -ListAvailable | foreach {"`n## Module name: $_ `n"; gcm -Module $_.name -CommandType $type | select name; "`r`n" } }
# -CommandType {Alias | Function | Filter | Cmdlet | ExternalScript | Application | Script | Workflow | Configuration | All}
# Get-Module -ListAvailable | foreach { foreach ($i in gcm -Module $_.name -CommandType $type | select name } }
# $types = @("Alias", "Function", "Filter", "Cmdlet", "ExternalScript", "Application", "Script", "Workflow", "Configuration")
function Get-Aliases { Get-CommandTypes Alias } ; Set-Alias aliases Get-Aliases
function Get-Functions { Get-CommandTypes Function } ; Set-Alias functions Get-Functions
function Get-Filters { Get-CommandTypes Filter } ; Set-Alias filters Get-Filters




# function Get-Functions { Get-Module -ListAvailable | foreach {"`n## Module name: $_ `n"; gcm -Module $_.name -CommandType function | select name; "`r`n" } }
# function Get-Cmdlets { Get-Module -ListAvailable | foreach {"`n## Module name: $_`n"; gcm -Module $_.name -CommandType cmdlet | select name; "`r`n" } }
Set-Alias cmdlets Get-Cmdlets 

# function Get-Aliases { Get-Module -ListAvailable | foreach {"`n## Module name: $_`n"; gcm -Module $_.name -CommandType alias | select name; "`r`n" } }
Set-Alias aliases Get-Aliases

function Get-Verbs { $out = ""; foreach ($i in ((get-verb).verb | sort)) { $out = "$out, $i" } ; "`n:: Registered PowerShell Verbs:`n" ; Write-Wrap $out.trimstart(", ") ; "" }
Set-Alias verbs Get-Verbs
# verbs finds all verb types available and sorts them. Using verbs outside of this range will result in 

function ver {
    $PSVersionTable
    ""
    echo "Win32_OperatingSystem : $((Get-WmiObject -class win32_operatingsystem).Version)"
    ""
    winver.exe
    $check = read-host "Check for Microsoft Office licenses (about 5 sec) (default is n) (y/n)? "
    if ($check -eq 'y') {
        "`nGet-WmiObject win32_product | where { `$_.Name -like ""*Office*"" } | select Name,Version,InstallDate | ft"   # Need to use '' due to the "*" wildcard
        Get-WmiObject win32_product | where { $_.Name -like "*Office*" } | select Name,Version | ft   #  -or $_.Name -Like "Microsoft Office Standard*" }
    }
    $check = read-host "Check *all* installed apps (about 5 sec) (default is n) (y/n)? "
    if ($check -eq 'y') {
        "`nGet-WmiObject win32_product | select Name,Version,InstallDate,Vendor | sort InstallDate | ft"
        Get-WmiObject win32_product | select Name,Version,InstallDate,Vendor | sort InstallDate | ft
    }
    # Find all mods, Get-Command -Module $module for all modules
    # display some kind of table showing all of this
    # if give a $module, then get as much info as possible for that module! is it installed, where is it located,
    # what files are in that folder, what versions are available? is there a manifest (.psf1), how many functions
    # are available.
}



# Quick Chocolatey Functions, c for Chocolatey, then 2-char shorthand for a command, then y for assume "yes" on all 
# There is little point in running these if not Administrator, so always offer to sudo them at runtime.
function ciny { 
    if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        choco install -y $args   # c(hoco) in(stall) -y
    }
    else {
        Write-Host "User is not elevated, so Chocolatey operation may fail." -b black -f red
        $answer = read-host "Would you like to gsudo this operation to run as Adminsitrator (y/n)? "
        if ($answer -eq 'y' -or $intput -eq '') {
            gsudo choco install -y $args   # c(hoco) in(stall) -y
        }
    }
}

function cuny {
    if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        choco uninstall -y $args   # c(hoco) un(install) -y
    }
    else {
        Write-Host "User is not elevated, so Chocolatey operation may fail." -b black -f red
        $answer = read-host "Would you like to gsudo this operation to run as Adminsitrator (y/n)? "
        if ($answer -eq 'y' -or $intput -eq '') {
            gsudo choco uninstall -y $args   # c(hoco) un(install) -y
        }
    }
} 
function cupy { choco upgrade -y all }       # cup (choco upgrade) all -y

# function g { } # test for existence of gsudo, if not there, choco install -y it and then use it with the command given

function Get-EventErrorSummary ($daysback) {
    Get-EventLog -LogName 'Application' -EntryType Error -After ((Get-Date).Date.AddDays(-$daysback)) |
      ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name LogDay -Value $_.TimeGenerated.ToString("yyyyMMdd") -PassThru } |
      Group-Object LogDay | Select-Object @{N='LogDay';E=    {[int]$_.Name}},Count | Sort-Object LogDay |
      Format-Table -Auto
}

function Get-EventTesting {
    Get-WinEvent -ListLog * -EA silentlycontinue | where-object { $_.recordcount -AND $_.lastwritetime -gt [datetime]::today} | ForEach-Object { Get-WinEvent -LogName $_.logname -MaxEvents 3 }
    # Work on this, break down Events by type, work on functions to extract generic information
    # Need to expand this a lot to try and pull some generic and useful predefined EventLog things like:
    # - Show all logons
    # - Show all times that a given app was started and stopped
    # - Show just all Errors in a given time perior (or Warnings or Information etc)
    # - Show all Security alerts.
    # - Also map each of the above with information on how to achieve the same from within the GUI
    # https://devblogs.microsoft.com/scripting/use-powershell-to-query-all-event-logs-for-recent-events/
}

# Set-Alias Get-Errors Get-EventErrorSummary
# Start-->Run-->Eventvwr-->Windows logs-->Security. Filter by 'Task Category = Logoff'
# https://stackoverflow.com/questions/32177668/find-user-disconnection-time-in-rdp-session-windows-server-2012



# ToDo: Remember the -WhatIf Common Parameter
# Rename-Item -NewName  {$_.Name -replace 'zip','OLD'} -WhatIf
# https://learn-powershell.net/2014/10/11/using-a-scriptblock-parameter-with-a-powershell-function/



# https://community.spiceworks.com/topic/398280-powershell-get-the-size-of-several-subfolders

# Function Get-FolderSize {
# 
# [cmdletbinding()]
# Param ( $folder = $(Throw "no folder name specified"))
# 
# # calculate folder size and recurse as needed
# $size = 0
# Foreach ($file in $(ls $folder -recurse)){
#  If (-not ($file.psiscontainer)) {
#     $size += $file.length
#     }
# }
# 
# # return the value and go back to caller
# return $size
# 
# }
# 
# get-foldersize C:\foo\bin


# Function Get-FolderSizes
# {   # requires -version 3.0
#     Param (
#         [Parameter(Mandatory=$true)]
#         [string]$Path
#     )
# 
#     If (Test-Path $Path -PathType Container)
#     {   ForEach ($Folder in (Get-ChildItem $Path -Directory -Recurse))
#         {   [PSCustomObject]@{
#                 Folder = $Folder.FullName
#                 Size = (Get-ChildItem $Folder.FullName -RecurseSize| Measure-Object Length -Sum).Sum
#             }
#         }
#     }
# }
# 
# Get-FolderSizes -Path c:\Dropbox\test



#Adapted from https://gist.github.com/altrive/5329377
#Based on <http://gallery.technet.microsoft.com/scriptcenter/Get-PendingReboot-Query-bdb79542>
function Test-PendingReboot {
    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { return $true }
    if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { return $true }
    if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA Ignore) { return $true }
    try { 
        $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
        $status = $util.DetermineIfRebootPending()
        if(($status -ne $null) -and $status.RebootPending) { return $true }
    }
    catch {}
    return $false
}

function Get-IPAddress {
    # Note to change all networks to Private (required for WinRM): Get-NetAdapter, Get-NetConnectionProfile
    # Set-NetConnectionProfile -Name "Brink-Router3" -NetworkCategory Private
    Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPEnabled -eq $true } |   # add two new properties for IPv4 and IPv6 at the end
    Select-Object -Property Description, MacAddress, IPAddress, IPAddressV4, IPAddressV6 |
    ForEach-Object {
        # add IP addresses that match the filter to the new properties
        $_.IPAddressV4 = $_.IPAddress | Where-Object { $_ -like '*.*.*.*' }
        $_.IPAddressV6 = $_.IPAddress | Where-Object { $_ -notlike '*.*.*.*' }
        # return the object
        $_
    } | Select-Object -Property Description, MacAddress, IPAddressV4, IPAddressV6    # remove the property that holds all IP addresses
}
Set-Alias ip Get-IPAddress
# function ipv4 { ipconfig | where { $_ -match "^Ethernet|^Wireless|IPv4" } } ; Set-Alias ip4 ipv4
# function ipv6 { ipconfig | where { $_ -match "^Ethernet|^Wireless|IPv6" } } ; Set-Alias ip6 ipv6
# function ip { ipconfig | where { $_ -match "^Ethernet|^Wireless|IPv4|IPv6" } }

function Enable-PSColor {    
    # Enable colour directory listings
    ""
    if ( (!(Test-Path "C:\Program Files\WindowsPowerShell\Modules\PSColor")) -and (!(Test-Path "C:\Users\$env:USERNAME\Documents\WindowsPowerShell\Modules\PSColor")) ) { 
        "Installing PSColor Module ..."
        Install-Module PSColor
    }
    if ( (Test-Path "C:\Program Files\WindowsPowerShell\Modules\PSColor") -or (Test-Path "C:\Users\$env:USERNAME\Documents\WindowsPowerShell\Modules\PSColor") ) { 
        "Importing PSColor Module ..."
        Import-Module PSColor
    }
    if (Get-Module -All PSColor) { 
        dir
    } else {
        "PSColor failed to install or import"
    }
}

Set-Alias pscolor Enable-PSColor
Set-Alias gcicolor Enable-PSColor
Set-Alias dircolor Enable-PSColor
Set-Alias lscolor Enable-PSColor

function Test-Colors {
    ""
    "To view all Console Colors:   [System.Enum]::GetValues('ConsoleColor')"
    ""
    "Test Color Pallete:"
    [Console]::ResetColor()
    ""
    "Running 'ConsoleArt' script imported by 'BeginSystemConfig.ps1' to the default"
    "PowerShell script repository at:   C:\Users\$env:USERNAME\Documents\WindowsPowerShell\Scripts"
    ""

    # $HomeFix = $HOME
    # $HomeLeaf = split-path $HOME -leaf   # Just get the correct username in spite of any changes to username (as on Edwin's system where username -ne foldername)
    # if ($HomeFix -like "\\*") { $HomeFix = "C:\Users\$(Split-Path $HOME -Leaf)" }
    # # The default Modules and Scripts paths are not created by default in Windows
    # if (!(Test-Path $HomeFix)) { md $HomeFix -Force -EA silent | Out-Null }
    # if (!(Test-Path "$HomeFix\Documents\WindowsPowerShell\Modules")) { md "$HomeFix\Documents\WindowsPowerShell\Modules" -Force -EA silent | Out-Null }
    # if (!(Test-Path "$HomeFix\Documents\WindowsPowerShell\Scripts")) { md "$HomeFix\Documents\WindowsPowerShell\Scripts" -Force -EA silent | Out-Null }
    # $CustomToolsPath = "$HomeFix\Documents\WindowsPowerShell\Modules\Custom-Tools\Custom-Tools.psm1"
    # $UserModulesPath = "$HomeFix\Documents\WindowsPowerShell\Modules"   # $UserModulesPath = "C:\Users\$HomeLeaf\Documents\WindowsPowerShell\Modules"
    # $UserScriptsPath = "$HomeFix\Documents\WindowsPowerShell\Scripts"
    # $AdminModulesPath = "C:\Program Files\WindowsPowerShell\Modules"

    function Download-Script ($url) {
        $FileName = ($url -split "/")[-1]   # Could also use:  $url -split "/" | select -last 1   # 'hi there, how are you' -split '\s+' | select -last 1
        $OutPath = Join-Path $UserScriptsPath $FileName 
        Write-Host "Downloading  $FileName to $OutPath ..."
        try { (New-Object System.Net.WebClient).DownloadString($url) | Out-File $OutPath }
        catch { "Could not download $FileName ..." }
    }

    Download-Script 'https://gist.github.com/shanenin/f164c483db513b88ce91/raw'
    if (Test-Path "$UserScriptsPath\raw") { Move-Item "$UserScriptsPath\raw" "$UserScriptsPath\ConsoleArt.ps1" -Force }

    echo $UserScriptsPath
    echo "Run ConsoleArt script showing a face drawn in the console and a staggered text...`n`n"
    ConsoleArt

    foreach($color1 in (0..15)) {
        Write-Host "    " -NoNewline
        foreach($color2 in (0..15)) { Write-Host -ForegroundColor ([ConsoleColor]$color1) -BackgroundColor ([ConsoleColor]$color2) -Object "X" -NoNewline } ; ""
    }

    $colors = [enum]::GetValues([System.ConsoleColor])
    Foreach ($bgcolor in $colors){
        Foreach ($fgcolor in $colors) {
            Write-Host "$fgcolor|"  -ForegroundColor $fgcolor -BackgroundColor $bgcolor -NoNewLine
        }
        Write-Host " on $bgcolor"
    }

    # $host.ui.rawui.ForegroundColor = <ConsoleColor>
    # $host.ui.rawui.BackgroundColor = <ConsoleColor>
    # $Host.PrivateData.ErrorForegroundColor = <ConsoleColor>
    # $Host.PrivateData.ErrorBackgroundColor = <ConsoleColor>
    # $Host.PrivateData.WarningForegroundColor = <ConsoleColor>
    # $Host.PrivateData.WarningBackgroundColor = <ConsoleColor>
    # $Host.PrivateData.DebugForegroundColor = <ConsoleColor>
    # $Host.PrivateData.DebugBackgroundColor = <ConsoleColor>
    # $Host.PrivateData.VerboseForegroundColor = <ConsoleColor>
    # $Host.PrivateData.VerboseBackgroundColor = <ConsoleColor>
    # $Host.PrivateData.ProgressForegroundColor = <ConsoleColor>
    # $Host.PrivateData.ProgressBackgroundColor = <ConsoleColor>

    # https://www.delftstack.com/howto/powershell/change-colors-in-powershell/
    # $host.PrivateData.ErrorBackgroundColor = "White"

    # https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/using-colors-in-powershell-console
    
    # https://www.commandline.ninja/easily-display-powershell-console-colors/
    $List = [enum]::GetValues([System.ConsoleColor]) 
    
    ForEach ($Color in $List){
        Write-Host "      $Color" -ForegroundColor $Color -NonewLine
        Write-Host "" 
        
    } #end foreground color ForEach loop

    ForEach ($Color in $List){
        Write-Host "                   " -backgroundColor $Color -noNewLine
        Write-Host "   $Color"
                
    } #end background color ForEach loop

    # https://stackoverflow.com/questions/64203354/set-the-text-color-in-powershell
    # https://stackoverflow.com/questions/36116326/programmatically-change-powershells-16-default-console-colours
    # https://github.com/lukesampson/concfg
    # https://stackoverflow.com/questions/70010554/changing-verbose-colors-in-powershell-7-2
    # https://stackoverflow.com/questions/51001708/background-color-and-console-text-color

    # https://stackoverflow.com/questions/16280402/setting-powershell-colors-with-hex-values-in-profile-script
    # cd hkcu:/console
    # $0 = '%systemroot%_system32_windowspowershell_v1.0_powershell.exe'
    # ni $0 -f
    # sp $0 ColorTable00 0x00562401
    # sp $0 ColorTable07 0x00f0edee
    
    # https://jdhitsolutions.com/blog/powershell/7753/powershell-color-combos/
        
    # https://4sysops.com/wiki/change-powershell-console-syntax-highlighting-colors-of-psreadline/
    # Get-PSReadlineOption  # list all.  (alias: just 'psreadlineoption')
    # Set-PSReadLineOption -Colors @{ "Command"="White" }
    # Set-PSReadLineOption -Colors @{ "Operator"="DarkBlue" }
    # Set-PSReadLineOption -Colors @{ "String"="Yellow" }
    # Set-PSReadLineOption -Colors @{ "Parameter"="Blue" }
    # Set-PSReadLineOption -Colors @{ "Comment"="Gray" }
    # # which syntax I found here:
    # get-help Set-PSReadLineOption -examples
    # Get-PSReadLineOption | out-string -stream | sls "char"
    # [System.Enum]::getvalues([System.ConsoleColor])

    # I have used all these, many are most likely duplicate colors, but they all work
    # $Colors = @()
    # 
    # $Colors += "AliceBlue"
    # $Colors += "AntiqueWhite"
    # $Colors += "Aqua"
    # $Colors += "Aquamarine"
    # $Colors += "Azure"
    # $Colors += "Beige"
    # $Colors += "Bisque"
    # $Colors += "Black"
    # $Colors += "BlanchedAlmond"
    # $Colors += "Blue"
    # $Colors += "BlueViolet"
    # $Colors += "Brown"
    # $Colors += "BurlyWood"
    # $Colors += "CadetBlue"
    # $Colors += "Chartreuse"
    # $Colors += "Chocolate"
    # $Colors += "Coral"
    # $Colors += "CornflowerBlue"
    # $Colors += "Cornsilk"
    # $Colors += "Crimson"
    # $Colors += "Cyan"
    # $Colors += "DarkBlue"
    # $Colors += "DarkCyan"
    # $Colors += "DarkGoldenrod"
    # $Colors += "DarkGray"
    # $Colors += "DarkGreen"
    # $Colors += "DarkKhaki"
    # $Colors += "DarkMagenta"
    # $Colors += "DarkOliveGreen"
    # $Colors += "DarkOrange"
    # $Colors += "DarkOrchid"
    # $Colors += "DarkRed"
    # $Colors += "DarkSalmon"
    # $Colors += "DarkSeaGreen"
    # $Colors += "DarkSlateBlue"
    # $Colors += "DarkSlateGray"
    # $Colors += "DarkTurquoise"
    # $Colors += "DarkViolet"
    # $Colors += "DeepPink"
    # $Colors += "DeepSkyBlue"
    # $Colors += "DimGray"
    # $Colors += "DodgerBlue"
    # $Colors += "Firebrick"
    # $Colors += "FloralWhite"
    # $Colors += "ForestGreen"
    # $Colors += "Fuchsia"
    # $Colors += "Gainsboro"
    # $Colors += "GhostWhite"
    # $Colors += "Gold"
    # $Colors += "Goldenrod"
    # $Colors += "Gray"
    # $Colors += "Green"
    # $Colors += "GreenYellow"
    # $Colors += "Honeydew"
    # $Colors += "HotPink"
    # $Colors += "IndianRed"
    # $Colors += "Indigo"
    # $Colors += "Ivory"
    # $Colors += "Khaki"
    # $Colors += "Lavender"
    # $Colors += "LavenderBlush"
    # $Colors += "LawnGreen"
    # $Colors += "LemonChiffon"
    # $Colors += "LightBlue"
    # $Colors += "LightCoral"
    # $Colors += "LightCyan"
    # $Colors += "LightGoldenrodYellow"
    # $Colors += "LightGray"
    # $Colors += "LightGreen"
    # $Colors += "LightPink"
    # $Colors += "LightSalmon"
    # $Colors += "LightSeaGreen"
    # $Colors += "LightSkyBlue"
    # $Colors += "LightSlateGray"
    # $Colors += "LightSteelBlue"
    # $Colors += "LightYellow"
    # $Colors += "Lime"
    # $Colors += "LimeGreen"
    # $Colors += "Linen"
    # $Colors += "Magenta"
    # $Colors += "Maroon"
    # $Colors += "MediumAquamarine"
    # $Colors += "MediumBlue"
    # $Colors += "MediumOrchid"
    # $Colors += "MediumPurple"
    # $Colors += "MediumSeaGreen"
    # $Colors += "MediumSlateBlue"
    # $Colors += "MediumSpringGreen"
    # $Colors += "MediumTurquoise"
    # $Colors += "MediumVioletRed"
    # $Colors += "MidnightBlue"
    # $Colors += "MintCream"
    # $Colors += "MistyRose"
    # $Colors += "Moccasin"
    # $Colors += "NavajoWhite"
    # $Colors += "Navy"
    # $Colors += "OldLace"
    # $Colors += "Olive"
    # $Colors += "OliveDrab"
    # $Colors += "Orange"
    # $Colors += "OrangeRed"
    # $Colors += "Orchid"
    # $Colors += "PaleGoldenrod"
    # $Colors += "PaleGreen"
    # $Colors += "PaleTurquoise"
    # $Colors += "PaleVioletRed"
    # $Colors += "PapayaWhip"
    # $Colors += "PeachPuff"
    # $Colors += "Peru"
    # $Colors += "Pink"
    # $Colors += "Plum"
    # $Colors += "PowderBlue"
    # $Colors += "Purple"
    # $Colors += "Red"
    # $Colors += "RosyBrown"
    # $Colors += "RoyalBlue"
    # $Colors += "SaddleBrown"
    # $Colors += "Salmon"
    # $Colors += "SandyBrown"
    # $Colors += "SeaGreen"
    # $Colors += "SeaShell"
    # $Colors += "Sienna"
    # $Colors += "Silver"
    # $Colors += "SkyBlue"
    # $Colors += "SlateBlue"
    # $Colors += "SlateGray"
    # $Colors += "Snow"
    # $Colors += "SpringGreen"
    # $Colors += "SteelBlue"
    # $Colors += "Tan"
    # $Colors += "Teal"
    # $Colors += "Thistle"
    # $Colors += "Tomato"
    # $Colors += "Turquoise"
    # $Colors += "Violet"
    # $Colors += "Wheat"
    # $Colors += "White"
    # $Colors += "WhiteSmoke"
    # $Colors += "Yellow"
    # $Colors += "YellowGreen"
}

function Enable-DirFriendlySizes {
    # After running this, dir/ls/gci will show friendly sizes (but also sortable etc)
    $file = '{0}myTypes.ps1xml' -f ([System.IO.Path]::GetTempPath()) 
    $data = Get-Content -Path $PSHOME\FileSystem.format.ps1xml
    # Note: Change $N=1 to $N=2 if want 2 decimal places in output
    $data -replace '<PropertyName>Length</PropertyName>', @'
<ScriptBlock>
if($$_ -is [System.IO.FileInfo]) {
    $this=$$_.Length; $sizes='Bytes,KB,MB,GB,TB,PB,EB,ZB' -split ','
    for($i=0; ($this -ge 1kb) -and ($i -lt $sizes.Count); $i++) {$this/=1kb}
    $N=1; if($i -eq 0) {$N=0}
    "{0:N$($N)} {1}" -f $this, $sizes[$i]
} else { $null }
</ScriptBlock>
'@ | Set-Content -Path $file
    Update-FormatData -PrependPath $file    
    # https://martin77s.wordpress.com/2017/05/20/display-friendly-file-sizes-in-powershell/
    # dir | Select-Object -Property Mode, LastWriteTime, @{N='SizeInKb';E={[double]('{0:N2}' -f ($_.Length/1kb))}}, Name | Sort-Object -Property SizeInKb
}

function dirfriendly { Enable-DirFriendlySizes ; dir ; "`nThe Enable-DirFriendlySizes function has been enabled for this session.`nFile sizes will be shown in human readable format for dir/ls/gci.`nA new shell will disable this option."  }
function dircolor { Enable-DirColors ; "`nThe Enable-DirColors function has been enabled for this session.`nFiles will be coloured according to their extension type for dir/ls/gci.`nA new shell will disable this option." }

Function Get-FileDialog($InitialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    #$OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

# edit: Notice the error action if the operation is canceled. That might be useful for your script.
Function Get-FolderBrowserDialog ( [string]$Description = "Select Folder", [string]$RootFolder = "Desktop" ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = $RootFolder
    $objForm.Description = $Description
    $Show = $objForm.ShowDialog()
    if ($Show -eq "OK") { return $objForm.SelectedPath }
    else { Write-Error "Operation cancelled by user." }
}

# https://docs.microsoft.com/en-us/powershell/scripting/samples/multiple-selection-list-boxes?view=powershell-6
# In the function, can either "return $x" or in the function change the $x declaration to "$xglobal:$x" to make it visible outside of the function
function ShowDialog( [string]$file ) { 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
   
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Choice-Box"
    $objForm.Size = New-Object System.Drawing.Size(300,200) 
    $objForm.StartPosition = "CenterScreen"
   
    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown( 
        {
            if ($_.KeyCode -eq "Enter") {
                $x = $objListBox.SelectedItem; $objForm.Close()
            }
        }
    )
     
    $objForm.Add_KeyDown(
        {
            if ($_.KeyCode -eq "Escape") {
                $objForm.Close()
            }
        }
    )
   
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click( { $global:x = $objListBox.SelectedItem; $objForm.Close() } )
    $objForm.Controls.Add($OKButton)
   
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click( { $objForm.Close() } )
    $objForm.Controls.Add($CancelButton)
   
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(280,20) 
    $objLabel.Text = "Please choose any of the below :"
    $objForm.Controls.Add($objLabel) 
   
    $objListBox = New-Object System.Windows.Forms.ListBox 
    $objListBox.Location = New-Object System.Drawing.Size(10,40) 
    $objListBox.Size = New-Object System.Drawing.Size(260,20) 
    $objListBox.Height = 80
   
    $items = gc $file | where { $_ -ne "" }
    
    foreach ($item in $items) {
        [void] $objListBox.Items.Add($item)     
    }
   
    $objForm.Controls.Add($objListBox) 
   
    $objForm.Topmost = $True
   
    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    return $x
}

# Handle scriptname and 
# https://stackoverflow.com/questions/817198/how-can-i-get-the-current-powershell-executing-file
# $scriptFull = $script:MyInvocation.MyCommand.Definition
# $scriptPath = $script:MyInvocation.MyCommand.Path
# $scriptName = $script:MyInvocation.MyCommand.Name
# $scriptPathOnly = $scriptPath.Split($scriptName)[0]
# $scriptPathNoExt = $scriptName.Split(".ps1")[0]
# $scriptLog = $scriptPathNoExt + ".log"   # log file with same name as script
# $scriptLog = $scriptName.Replace(".ps1",".log")
# $scriptInput = $scriptName.Replace(".ps1",".txt")
# echo $scriptFull $scriptName $scriptPathNoExt $scriptPathOnly $scriptInput
# 
# $OutputEncoding = ShowDialog $scriptInput
# write-host $out



# Very quickly resolve the drives
function Get-AllDrives {
    Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType = '3'" | 
        Select-Object -Property DeviceID, DriveType, VolumeName, 
        @{L="Capacity GB";E={"{0:N2}" -f ($_.Size/1GB)}},
        @{L='FreeSpace GB';E={"{0:N2}" -f ($_.FreeSpace /1GB)}}
}

# Very quick setup for basic shares, mountpoints, etc
function Update-SharingAllDrives {
    foreach ($i in $(Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType = '3'")) {   # | Select-Object -Property DeviceID, VolumeName)
        echo "net share Drive-$($i.DeviceID)=$($i.DeviceID)\ /grant:Everyone,Full /unlimited /remark:`"Full Access, $($i.VolumeName)`""
        $i.GetRelationships() | Select-Object -Property __RELPATH
        ""
    }
}

# https://techibee.com/powershell/powershell-how-to-get-list-of-mapped-drives/1156
# https://mcpmag.com/articles/2018/01/26/view-drive-information-with-powershell.aspx
function Get-RemoteDrives ($computernames) {
    # $computernames = Get-Content 'I:\NodeList\SNL.txt'
    $CSVpath = "C:\0\RemoveDrives.csv"
    Remove-Item $CSVpath -Force -EA Silent  
    $Report = @()

    foreach ($computer in $computernames) {
        Write-host $computer
        $colDrives = Get-WmiObject Win32_MappedLogicalDisk -ComputerName $computer

        foreach ($objDrive in $colDrives) {
            # For each mapped drive - build a hash containing information
            $hash = @{
                ComputerName = $computer
                MappedLocation = $objDrive.ProviderName
                DriveLetter = $objDrive.DeviceId
            }
            # Add the hash to a new object
            $objDriveInfo = New-Object PSObject -Property $hash
            # Store our new object within the report array
            $Report += $objDriveInfo
        }
        # Export our report array to CSV and store as our dynamic file name
        $Report | Export-Csv -NoType $CSVpath #$filenamestring
    }
}

# An easy way to get the functionality of a list box is to use Out-GridView. For example:
# 
# 'one','two','three','four' | Out-GridView -OutputMode Single
# The user can select an item, and it will be returned in the pipeline.
# 
# You can also use -OutputMode Multiple for multi-selection.
# 
# Another example:
# 
# get-process | Out-GridView -OutputMode Multiple
# This returns the selected object:
# 
# Handles  NPM(K)    PM(K)      WS(K) VM(M)   CPU(s)     Id ProcessName
# -------  ------    -----      ----- -----   ------     -- -----------
#     396      24     5008       8964   113     8,88   3156 AuthManSvr



# https://binarynature.blogspot.com/2010/04/powershell-version-of-df-command.html
# PowerShell equivalent of the df command
# Get-DiskFree | Get-Member
# 'db01','sp01' | Get-DiskFree -Credential $cred -Format | ft -GroupBy Name -auto  
# Name Vol Size  Used  Avail Use% FS   Type
# ---- --- ----  ----  ----- ---- --   ----
# DB01 C:  39.9G 15.6G 24.3G   39 NTFS Local Fixed Disk
# DB01 D:  4.1G  4.1G  0B     100 CDFS CD-ROM Disc
### Low Disk Space: just get list of servers in AD with disk space below 20% for C: volume?
# Import-Module ActiveDirectory
# $servers = Get-ADComputer -Filter { OperatingSystem -like '*win*server*' } | Select-Object -ExpandProperty Name
# Get-DiskFree -cn $servers | Where-Object { ($_.Volume -eq 'C:') -and ($_.Available / $_.Size) -lt .20 } | Select-Object Computer
### Out-GridView: filter on drives of four servers and have the output displayed in an interactive table.
# $cred = Get-Credential 'example\administrator'
# $servers = 'dc01','db01','exch01','sp01'
# Get-DiskFree -Credential $cred -cn $servers -Format | ? { $_.Type -like '*fixed*' } | select * -ExcludeProperty Type | Out-GridView -Title 'Windows Servers Storage Statistics'
### Output to CSV: similar to the previous except we will also sort the disks by the percentage of usage. We've also decided to narrow the set of properties to name, volume, total size, and the percentage of the drive space currently being used.
# $cred = Get-Credential 'example\administrator'
# $servers = 'dc01','db01','exch01','sp01'
# Get-DiskFree -Credential $cred -cn $servers -Format | ? { $_.Type -like '*fixed*' } | sort 'Use%' -Descending | select -Property Name,Vol,Size,'Use%' | Export-Csv -Path $HOME\Documents\windows_servers_storage_stats.csv -NoTypeInformation
# https://www.computerperformance.co.uk/powershell/format-table/
function Get-DiskFree
{
    <#
    <#
    .SYNOPSIS
    Short description...
    Get-DiskFree function (Unix df equivalent, check all volumes for disk used and free)
    .DESCRIPTION
    Long description...
    Get-DiskFree function (Unix df equivalent, check all volumes for disk used and free)
    .PARAMETER Arguments
    Parameters ... stuff
    .EXAMPLE
    'sys01','sys01' | Get-DiskFree -Credential $cred -Format | ft -GroupBy Name -auto

    Low Disk Space: just get list of servers in AD with disk space below 20% for C: volume?
    Import-Module ActiveDirectory
    $servers = Get-ADComputer -Filter { OperatingSystem -like '*win*server*' } | Select-Object -ExpandProperty Name
    Get-DiskFree -cn $servers | Where-Object { ($_.Volume -eq 'C:') -and ($_.Available / $_.Size) -lt .20 } | Select-Object Computer
    
    Out-GridView: filter on drives of four servers and have the output displayed in an interactive table.
    $cred = Get-Credential 'example\administrator'
    $servers = 'dc01','db01','exch01','sp01'
    Get-DiskFree -Credential $cred -cn `$servers -Format | ? { $_.Type -like '*fixed*' } | select * -ExcludeProperty Type | Out-GridView -Title 'Windows Servers Storage Statistics'
    
    Output to CSV: similar to the previous except we will also sort the disks by the percentage of usage. We've also decided to narrow the set of properties to name, volume, total size, and the percentage of the drive space currently being used.
    $cred = Get-Credential 'example\administrator'
    $servers = 'dc01','db01','exch01','sp01'
    Get-DiskFree -Credential $cred -cn $servers -Format | ? { $_.Type -like '*fixed*' } | sort 'Use%' -Descending | select -Property Name,Vol,Size,'Use%' | Export-Csv -Path $HOME\Documents\windows_servers_storage_stats.csv -NoTypeInformation
    .NOTES
    Some notes ...
    
    Based on code posted by weestro at http://weestro.blogspot.com/
    http://weestro.blogspot.com/2009/08/sudo-for-powershell.html
    https://joeit.wordpress.com/
    .LINK
    https://joeit.wordpress.com/
    .LINK
    http://weestro.blogspot.com/2009/08/sudo-for-powershell.html
    #>

    # Get-DiskFree function (Unix df equivalent, check all volumes for disk used and free)"
    # 'sys01','sys01' | Get-DiskFree -Credential `$cred -Format | ft -GroupBy Name -auto"
    # ### Low Disk Space: just get list of servers in AD with disk space below 20% for C: volume?"
    # Import-Module ActiveDirectory"
    # `$servers = Get-ADComputer -Filter { OperatingSystem -like '*win*server*' } | Select-Object -ExpandProperty Name"
    # Get-DiskFree -cn `$servers | Where-Object { (`$_.Volume -eq 'C:') -and (`$_.Available / `$_.Size) -lt .20 } | Select-Object Computer"
    # ### Out-GridView: filter on drives of four servers and have the output displayed in an interactive table."
    # `$cred = Get-Credential 'example\administrator'"
    # `$servers = 'dc01','db01','exch01','sp01'"
    # Get-DiskFree -Credential `$cred -cn `$servers -Format | ? { `$_.Type -like '*fixed*' } | select * -ExcludeProperty Type | Out-GridView -Title 'Windows Servers Storage Statistics'"
    # ### Output to CSV: similar to the previous except we will also sort the disks by the percentage of usage. We've also decided to narrow the set of properties to name, volume, total size, and the percentage of the drive space currently being used."
    # `$cred = Get-Credential 'example\administrator'"
    # `$servers = 'dc01','db01','exch01','sp01'"
    # # Get-DiskFree -Credential `$cred -cn `$servers -Format | ? { `$_.Type -like '*fixed*' } | sort 'Use%' -Descending | select -Property Name,Vol,Size,'Use%' | Export-Csv -Path $HOME\Documents\windows_servers_storage_stats.csv -NoTypeInformation"
    [CmdletBinding()]param
    (
        [Parameter(Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias('hostname')]
        [Alias('cn')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        
        [Parameter(Position=1,
                   Mandatory=$false)]
        [Alias('runas')]
        [System.Management.Automation.Credential()]$Credential =
        [System.Management.Automation.PSCredential]::Empty,
        
        [Parameter(Position=2)]
        [switch]$Format
    )
    
    BEGIN
    {
        function Format-HumanReadable 
        {
            param ($size)
            switch ($size) 
            {
                {$_ -ge 1PB}{"{0:#.#'P'}" -f ($size / 1PB); break}
                {$_ -ge 1TB}{"{0:#.#'T'}" -f ($size / 1TB); break}
                {$_ -ge 1GB}{"{0:#.#'G'}" -f ($size / 1GB); break}
                {$_ -ge 1MB}{"{0:#.#'M'}" -f ($size / 1MB); break}
                {$_ -ge 1KB}{"{0:#'K'}" -f ($size / 1KB); break}
                default {"{0}" -f ($size) + "B"}
            }
        }
        
        $wmiq = 'SELECT * FROM Win32_LogicalDisk WHERE Size != Null AND DriveType >= 2'
    }
    
    PROCESS
    {
        foreach ($computer in $ComputerName)
        {
            try
            {
                if ($computer -eq $env:COMPUTERNAME)
                {
                    $disks = Get-WmiObject -Query $wmiq `
                             -ComputerName $computer -ErrorAction Stop
                }
                else
                {
                    $disks = Get-WmiObject -Query $wmiq `
                             -ComputerName $computer -Credential $Credential `
                             -ErrorAction Stop
                }
                
                if ($Format)
                {
                    # Create array for $disk objects and then populate
                    $diskarray = @()
                    $disks | ForEach-Object { $diskarray += $_ }
                    
                    $diskarray | Select-Object @{n='Name';e={$_.SystemName}}, 
                        @{n='Vol';e={$_.DeviceID}},
                        @{n='Size';e={Format-HumanReadable $_.Size}},
                        @{n='Used';e={Format-HumanReadable `
                        (($_.Size)-($_.FreeSpace))}},
                        @{n='Avail';e={Format-HumanReadable $_.FreeSpace}},
                        @{n='Use%';e={[int](((($_.Size)-($_.FreeSpace))`
                        /($_.Size) * 100))}},
                        @{n='FS';e={$_.FileSystem}},
                        @{n='Type';e={$_.Description}}
                }
                else 
                {
                    foreach ($disk in $disks)
                    {
                        $diskprops = @{'Volume'=$disk.DeviceID;
                                   'Size'=$disk.Size;
                                   'Used'=($disk.Size - $disk.FreeSpace);
                                   'Available'=$disk.FreeSpace;
                                   'FileSystem'=$disk.FileSystem;
                                   'Type'=$disk.Description
                                   'Computer'=$disk.SystemName;}
                    
                        # Create custom PS object and apply type
                        $diskobj = New-Object -TypeName PSObject `
                                   -Property $diskprops
                        $diskobj.PSObject.TypeNames.Insert(0,'BinaryNature.DiskFree')
                    
                        Write-Output $diskobj
                    }
                }
            }
            catch 
            {
                # Check for common DCOM errors and display "friendly" output
                switch ($_)
                {
                    { $_.Exception.ErrorCode -eq 0x800706ba } `
                        { $err = 'Unavailable (Host Offline or Firewall)'; 
                            break; }
                    { $_.CategoryInfo.Reason -eq 'UnauthorizedAccessException' } `
                        { $err = 'Access denied (Check User Permissions)'; 
                            break; }
                    default { $err = $_.Exception.Message }
                }
                Write-Warning "$computer - $err"
            } 
        }
    }  
    END {}
}
function df { Get-DiskFree -Format | Format-Table }
# { Get-DiskFree -Format | ft -GroupBy Name -auto }

# Older, simpler df version
# function df {
#     $colItems = Get-wmiObject -class "Win32_LogicalDisk" -namespace "root\CIMV2" -computername localhost
#     echo "DevID`t FSname`t Size GB`t FreeSpace GB`t Description"
# 
#     foreach ($objItem in $colItems) {
#         $DevID = $objItem.DeviceID
# 	      $FSname = $objItem.FileSystem
#         $size = ($objItem.Size / 1GB).ToString("f2")
# 	      $FreeSpace = ($objItem.FreeSpace / 1GB).ToString("f2")
# 	      $description = $objItem.Description
#         echo "$DevID`t $FSname`t $Size GB`t $FreeSpace GB`t $Description"
#     }	
# }



# https://ilovepowershell.com/2015/06/05/find-the-processes-using-the-most-cpu-on-a-computer-with-powershell/
# Get Highest CPU processes on local or remote systems. Use -Count to specify number of processes.
Function Get-HighestCPU {   
    <#
    .SYNOPSIS
        Retrieve processes that are utilizing the CPU on local or remote systems.
    
    .DESCRIPTION
        Uses WMI to retrieve process information from remote or local machines. You can specify to return X number of the top CPU consuming processes
        or to return all processes using more than a certain percentage of the CPU.
    .EXAMPLE
         Get-HighCPUProcess
        Returns the 3 highest CPU consuming processes on the local system.
    .EXAMPLE
         Get-HighCPUProcess -Count 1 -Computername AppServer01
        Returns the 1 highest CPU consuming processes on the remote system AppServer01.   
    .EXAMPLE
         Get-HighCPUProcess -MinPercent 15 -Computername "WebServer15","WebServer16"
        Returns all processes that are consuming more that 15% of the CPU process on the hosts webserver15 and webserver160
    #>
    
    [Cmdletbinding(DefaultParameterSetName="ByCount")]
    Param(
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias("PSComputername")]
        [string[]]$Computername = "localhost",
        
        [Parameter(ParameterSetName="ByPercent")]
        [ValidateRange(1,100)]
        [int]$MinPercent,
    
        [Parameter(ParameterSetName="ByCount")]
        [int]$Count = 3
    )
    
    Process {
        Foreach ($computer in $Computername){
        
            Write-Verbose "Retrieving processes from $computer"
            $wmiProcs = Get-WmiObject Win32_PerfFormattedData_PerfProc_Process -Filter "idProcess != 0" -ComputerName $Computername
        
            if ($PSCmdlet.ParameterSetName -eq "ByCount") {
                $wmiObjects = $wmiProcs | Sort PercentProcessorTime -Descending | Select -First $Count
            } elseif ($psCmdlet.ParameterSetName -eq "ByPercent") {
                $wmiObjects = $wmiProcs | Where {$_.PercentProcessorTime -ge $MinPercent} 
            } #end IF
    
            $wmiObjects | Foreach {
                $outObject = [PSCustomObject]@{
                    Computername = $computer
                    ProcessName = $_.name
                    Percent = $_.PercentProcessorTime
                    ID = $_.idProcess
                }
                $outObject
            } #End Foreach wmiObject
        } #End Foreach Computer
    }    
}
    
# Concise output: Computername, ProcessName, Percent (of CPU), and ID (Process ID)
# Accepts pipeline input by property value of "ComputerName" or "PSComputerName"
# Easy to use against multiple computers
# Flexible: I like the option to find the 3 highest CPU processes or all processes above 5 percent (or any percent)
# Accepts multiple computers through parameter (as an array) or through the pipeline



# Note on declaring parameters with preset values and types
# param( [string]$dir = "C:\Windows", [int32]$size = 200000 )
#
# With the above params, the following will test files in arg[0]
# to see if they are bigger than the value of arg[1]
#
# $files = Get-ChildItem $args[0]
# foreach ($file in $files) { if ($file.length -gt $args[1]) { Write-Output $file }


# Alternative GetDateTimeFormats (there are 114 standard time format outputs in this)
# for ($i=1; $i -lt 114; $i++)  { Write-Host "$i :" (Get-Date).GetDateTimeFormats()[$i].ToString() }   # view all TimeFormat
# (Get-Date).GetDateTimeFormats()[77].ToString()   # 2019-11-19 21:03:20
#    -replace " ", "__" -replace ":", "-"          # 2019-11-19__21-03-20
# (Get-Date).GetDateTimeFormats()[57].ToString()   # 2019-11-19 21:03
#    -replace " ", "__" -replace ":", "-"          # 2019-11-19__21-03
# (Get-Date).GetDateTimeFormats()[88].ToString()   # 21:03      -replace ":", "-" to make filename compatible
# (Get-Date).GetDateTimeFormats()[92].ToString()   # 21:03:14   -replace ":", "-" to make filename compatible



# https://www.experts-exchange.com/questions/26596458/Powershell-to-List-Folder-Sizes.html
# https://4sysops.com/archives/measure-object-computing-the-size-of-folders-and-files-in-powershell/
# gci -Dir -r | %{$_.FullName + "," + ((gci -File $_.FullName | measure Length -Sum).Sum) /1MB }
# You can just test $args variable or $args.count to see how many vars are passed to the script.
# Also, $args[0] -eq $null is different from $args[0] -eq 0 and from !$args[0]
# https://stackoverflow.com/questions/18607788/how-to-distinguish-between-empty-argument-and-zero-value-argument-in-powershell

# Get-ChildItem -Path $env:windir -Filter *.dll -Recurse
# 'splat' format:
# $myargs = @{
#   Path = "$env:windir"
#   Filter = '*.dll
#   Recurse = $true
# }
# Get-ChildItem @myargs

# $start = "C:\Program Files"   # Default location if no argument given
# if ($args.Count -eq 1) { $start = $args[0] }

function Get-SizeRecurse ( [string]$path, [long]$d ) {

    if (-not (Test-Path $path)) {
        Write-Error "$path is an invalid path."
        return $false
    }  
    
    try {
        $files = @(Get-ChildItem -Path $path -ErrorAction Ignore)
        
        $countfiles = 0 ; $countdirs = 0 ; $total = 0
        foreach ($file in $files) {

            if ($file.GetType().FullName -eq 'System.IO.FileInfo') {
                $total += $file.Length
                $countfiles++

            }
            elseif ($file.GetType().FullName -eq 'System.IO.DirectoryInfo') {
                $total += Get-SizeRecurse $file.FullName ($d + 1)
                $countdirs++
            }
        }
    } catch {
        Write-Host "Cannot access $file"
    }
    
    if ($d -lt 2)
    {
         $size = "{0:N2}" -f ($total / 1MB)
         Write-Host ('"' + $path + '","' + $size +' MB","files:' + $countfiles + '","dirs:' + $countdirs + '"')
    }
    
    [long] $total
}

function Disable-Cortana ([switch]$ExplorerRestart) {
    # http://www.alltechflix.com/disable-uninstall-cortana-windows-10/
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    if (!(Test-Path -Path $path)) { New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "Windows Search" }
    Set-ItemProperty -Path $path -Name "AllowCortana" -Value 0
    if ($ExplorerRestart -eq $true) { Stop-Process -Name Explorer }
}

function Update-RegItem ($RegPath, $Item, $Value) {
    $Path = Split-Path $RegPath
    $Name = Split-Path $RegPath -Leaf  # For HKLM:\SOFTWARE\Microsoft\Windows, -Leaf would be "Windows"
    if (!(Test-Path -Path $RegPath)) { New-Item -Path $Path -Name $Name -Force }
    if (!(Test-Path -Path "$RegPath\$Item")) { New-ItemProperty -Path $RegPath -Name $Item -Value $Value -Force }
    else { Set-ItemProperty -Path $RegPath -Name $Item -Value $Value -Force }
}
# Update-Registry "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" "AllowCortana" 0
# Update-Registry "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" "SearchboxTaskbarMode" 0

function Restart-Explorer
{
    Param([switch] $SuppressReOpen)
    # Cleanly restart Explorer.exe, but remember and reopening all currently open Explorer Windows
    # Restart-Explorer -SuppressReOpen to skip reopening the existing windows.

    # Gather up the currently open windows, so we can re-spawn them.
    $x = New-Object -ComObject "Shell.Application"
    $count = $x.Windows().Count
    $windows = @();
    $explorerPath = [System.IO.Path]::Combine($env:windir, "explorer.exe");
    for ($i=0; $i -lt $count; $i++)
    {
        # The location URL contains the Path that the explorer window is currently open to
        $url = $x.Windows().Item($i).LocationURL;

        $fullName = $x.Windows().Item($i).FullName;

        # This also catches IE windows, so I only add items where the full name is %WINDIR%\explorer.exe 
        if ($fullName.Equals($explorerPath))
        {
            $windows += $url
        }
    }

    Stop-Process -ProcessName explorer

    if (!$SuppressReOpen)
    {
        foreach ($window In $windows){

            if ([string]::IsNullOrEmpty($window)){
                Start-Process $explorerPath
            }
            else
            {
                Start-Process $window
            }
        }
    }
}
# Set-Alias re Restart-Explorer
# Export-ModuleMember -function Restart-Explorer -Alias re


# $bytes = RecurseSize $start 0
# $bytes = RecurseSize "C:\Program Files (x86)" 0
# $bytes = RecurseSize "C:\Windows" 0

# $startfolder = "c:\test"
# $folders = get-childitem $startfolder | where{$_.PSiscontainer -eq "True"}
# foreach ($fol in $Folders) {
#     $colItems = (Get-ChildItem $fol.fullname | Measure-Object -property length -sum)
#     $size = "{0:N2}" -f ($colItems.sum / 1MB) + " MB"
#     write-host "$($fol.fullname), $size"
# }


# Need to make this generic: BackupFolder ($path), 
# Then make various functions, function BackupProfile { BackupFolder "$($env:USERPROFILE)" }, etc.
# Also some work on media type and zero-size backups ( /create )
#    robocopy C:\0 ..\0_Backup_Books_zero_size *.mobi *.epub *.pdf /s /purge /create /r:1 /w:1
#    robocopy C:\0 .\0\ /mir /r:1 /w:1
#    robocopy C:\Users\Boss\ .\Boss\ /mir /r:1 /w:1


function Backup-Profile {
    $now = Get-Date -format "yyyy-MM-dd__HH-mm"

    $message = "`nStart backup of $($env:USERPROFILE)?`nBackup Location will be D:\Backup\$($env:USERNAME)_$($now)`n" # Don't use Confirm-Choice(); keep this function self-contained
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes";
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No";
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no);

    # Here, a couple things should happen:
    # - List all of the previous backups that match the -Leaf of the folder being backed up for reference.
    # - Offer to overwrite the last backup and update the folder name to the current date-time stamp now.

    $answer = $host.ui.PromptForChoice("", $message, $choices, 0)   # Set to 0 to default to "yes" and 1 to default to "no"

    if ($answer -eq '0') {
        if (!(Test-Path "D:\Backup")) { New-Item -Type Directory "D:\Backup" -EA silent }
        if (Test-Path "D:\Backup") {
            # Add a timespan to the whole operation
            robocopy "$($env:USERPROFILE)" "D:\Backup\$($env:USERNAME)_$($now)" /mir /r:1 /w:1 /xjd /xf ntuser.dat*
            # Important switches:
            # /XJ :: eXclude Junction points and symbolic links (junctions / symbolic links normally included by default).
            # /FFT :: assume FAT File Times (2-second granularity).
            # /DST :: compensate for one-hour DST time differences.
            # /XJD :: eXclude Junction points and symbolic links for Directories.   <<< Quite important for home folder
            # /XJF :: eXclude symbolic links for Files.

            # By defalut, robocopy will make the destination folder 'hidden'
            # https://stackoverflow.com/questions/464777/how-do-i-change-a-files-attribute-using-powershell
            function Get-FileAttribute {
                param($file,$attribute)
                $val = [System.IO.FileAttributes]$attribute
                if((gci $file -force).Attributes -band $val -eq $val){$true;} else { $false }
            } 
            
            function Set-FileAttribute {
                param($file,$attribute)
                $file =(gci $file -force)
                $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]$attribute).value__;
                if($?){$true} else {$false}
            }

            attrib -r -h -s -a "D:\Backup"
            # Note that you cannot change the "Compressed" attribute with the above.
            # If you want to set the compressed attribute on a folder using PowerShell you have to use the command-line tool compact:
            # compact /C /S c:\MyDirectory
        }
    }
}



# Create a zero-file-size backup of a folder for reference purposes
function Backup-ZeroSize ($source) {
    if ($source -eq $null) { "No folder specified." ; break }
    if ($source -eq ".") { $source = (Get-Location).Path }
   
    $now = Get-Date -format "yyyy-MM-dd__HH-mm"
    $extramsg = "`n"
    $destC = "C:\0\Backup" + "-" + ($source).replace(":\", "-").replace("\", "-") + "__" + $now + "_[zero-size]"
    $destD = "D:\Backup\$(hostname)" + "-" + ($source).replace(":\", "-").replace("\", "-") + "__" + $now + "_[zero-size]"
    if (Test-Path $destC) { $dest = $destC}
    if (Test-Path $destD) { $dest = $destD}

    if ($dest -like "$($source)*") { $extra = "/xd `"$dest`"" ; $extramsg = "`nExcluding destination folder as it is a subfolder of the source." }   # $extra: if $dest is a subfolder of $source, then exclude $dest: "/xd "$dest""
    if ($dest -eq $source) { "Destination is same as source" ; break }

    $message = "`nBackup Location will be $dest`nStart archive backup of $source (create only zero-size files for each file)?$extramsg" # Don't use Confirm-Choice(); keep this function self-contained
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes";
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No";
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no);
    $answer = $host.ui.PromptForChoice("", $message, $choices, 0)   # Set to 0 to default to "yes" and 1 to default to "no"
    if ($answer -eq '0') {
        if (!(Test-Path "D:\Backup")) { New-Item -Type Directory "D:\Backup" -EA silent }
        # Add a timespan to the whole operation        
        # Remember that if $source / $dest are objects, do not need " around them, but if string, need to surround with "
        $toRun = "robocopy `"$source`" `"$dest`" /mir /r:1 /w:1 /xjd /s /purge /create $extra"
        Write-Host "`n$($toRun)"

        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no);
        $answer = $host.ui.PromptForChoice("", "Start backup?", $choices, 0)   # Set to 0 to default to "yes" and 1 to default to "no"

        # https://stackoverflow.com/questions/6338015/how-do-you-execute-an-arbitrary-native-command-from-a-string
        if ($answer -eq '0') {
            iex $toRun
        }
        # /XJ :: eXclude Junction points and symbolic links. (normally included by default).
        # /FFT :: assume FAT File Times (2-second granularity).
        # /DST :: compensate for one-hour DST time differences.
        # /XJD :: eXclude Junction points and symbolic links for Directories.   <<< Quite important for home folder
        # /XJF :: eXclude symbolic links for Files.
    }
}

function hh {
    <#
    .SYNOPSIS
    http://woshub.com/powershell-commands-history/
        Get-PSReadlineOption
        Get-PSReadlineOption | select HistoryNoDuplicates, MaximumHistoryCount, HistorySearchCursorMovesToEnd, HistorySearchCaseSensitive, HistorySavePath, HistorySaveStyle
            HistoryNoDuplicates - determines whether the same commands have to be saved;
            MaximumHistoryCount - the maximum number of the stored commands (by default the last 4096 commands are saved);
            HistorySearchCursorMovesToEnd - determines whether you have to go to the end of the command when searching;
            HistorySearchCaseSensitive - determines whether search is case sensitive;
            HistorySavePath - shows the path to the file in which the command is stored;
            HistorySaveStyle - determines the peculiarities of saving commands:
                SaveIncrementally - the commands are saved after they are run (by default);
                SaveAtExit - the history is saved when the PowerShell console is closed;
                SaveNothing - disable saving command history.
    You can change the settings of PSReadLine module using Set-PSReadlineOption, for example:
        Set-PSReadlineOption -HistorySaveStyle SaveAtExit
    Remove-Item (Get-PSReadlineOption).HistorySavePath   # Delete history by deleting the history file (close the PowerShell window to complete history deletion)
    Set-PSReadlineOption -HistorySaveStyle SaveNothing   # To disable saving history
    function HistoryOff { Set-PSReadlineOption -HistorySaveStyle SaveNothing }
    function HistoryOn  { Set-PSReadlineOption -HistorySaveStyle SaveIncrementally }
    Remove-Item (Get-PSReadlineOption).HistorySavePath   # Delete history by deleting the history file (close the PowerShell window to complete history deletion)
#>
    Get-History | ft -Property *
    "Help files (stored separately for PowerShell, ISE, VSCode) are saved in:`n`$HOME\AppData\Roaming\Microsoft\Windows\PowerShell\PSReadline\ConsoleHost_history.txt`n"
    $System = get-wmiobject -class "Win32_ComputerSystem"
    $Mem = [math]::Ceiling($System.TotalPhysicalMemory / 1024 / 1024 / 1024)

    # $wmi = gwmi -class Win32_OperatingSystem -computer "."   # Removed this method as not CIM compliant
    # $LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
    # [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
    $BootUpTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $CurrentDate = Get-Date
    $Uptime = $CurrentDate - $BootUpTime
    $s = "" ; if ($Uptime.Days -ne 1) {$s = "s"}
    $uptime_string = "$($uptime.days) day$s $($uptime.hours) hr $($uptime.minutes) min $($uptime.seconds) sec"
    
    "Hostname (Domain): $($System.Name) ($($System.Domain)),   Make/Model: $($System.Manufacturer)/($($System.Model))"
    "PowerShell $($PSVersionTable.PSVersion),   Windows Version: $($PSVersionTable.BuildVersion),   Windows ReleaseId: $((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'ReleaseId').ReleaseId)"
    "Last Boot Time:  $([Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem | select 'LastBootUpTime').LastBootUpTime)),   Uptime: $uptime_string"
    $IPDefaultAddress = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].IPAddress[0]
    $IPDefaultGateway = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].DefaultIPGateway[0]
    "[Default IPAddress : $IPDefaultAddress / $IPDefaultGateway]"
    ""
}

function Fix-Sounds {
    "The default Windows sounds are annoying and jarring, ding.wav is much shorter and less annoying so replace various sounds by this."
    $toChange = @('.Default','SystemAsterisk','SystemExclamation','Notification.Default','SystemNotification','WindowsUAC','SystemHand')
    foreach ($c in $toChange) {
        Set-ItemProperty -Path "HKCU:\AppEvents\Schemes\Apps\.Default\$c\.Current\" -Name "(Default)" -Value "C:\WINDOWS\media\ding.wav"
    }
}

function Fix-LotsOfThings {
    @"
Stub function to collect many annoying things to fix.
Currently, this function does not do anything, just information, might build out to do stuff, and will rename if so...

VS Code and PSScriptAnalyzer are incredibly useful for correcting syntax errors, but flagging "gci / ls / dir" with a warning because they are aliases gets old fast.
Fix that by creating a white list https://github.com/PowerShell/PSScriptAnalyzer/blob/master/RuleDocumentation/AvoidUsingCmdletAliases.md
https://github.com/PowerShell/PSScriptAnalyzer/issues/214   # Some points here are fair, but not every script is intended as an "Enterprise Production Mission Critical" thing.
They go way way way over the top with the "never use aliases" mantra, and it's ridiculous when we use "-gt" as an abbreviation / common-English-language-alias for "greater than" 
and that's ok, even though, apparently, using "gci / ls / dir" will break the entire world?? Fuck that...
... details and automate that here ...
Open the settings file:
C:\Users\Boss\Documents\WindowsPowerShell\Modules\PSScriptAnalyzer\1.19.1\PSScriptAnalyzer.psd1
Add an entry for each item to be whitelisted
@{
    'Rules' = @{
        'PSAvoidUsingCmdletAliases' = @{
            'Whitelist' = @('cd', 'gci', 'ls', 'dir')
        }
    }
}
Actually, above is WRONG
Rules = @{
    PSAvoidUsingCmdletAliases = @{
        Whitelist = @('cd', 'gci', 'ls', 'dir')
    }
}

"@ | more
}

function Test-ModuleUpdate {
    # https://gist.github.com/jdhitsolutions/8a49a59c5dd19da9dde6051b3e58d2d0

    [cmdletbinding()] Param($mods)
    
    Write-Host "Getting installed modules" -ForegroundColor Yellow
    $modules = Get-Module -ListAvailable $mods
    
    #group to identify modules with multiple versions installed
    $g = $modules | group name -NoElement | where count -gt 1
    
    Write-Host "Filter to modules from the PSGallery" -ForegroundColor Yellow
    $gallery = $modules.where({$_.repositorysourcelocation})
    
    Write-Host "Comparing to online versions" -ForegroundColor Yellow
    foreach ($module in $gallery) {
    
         #find the current version in the gallery
         Try {
            $online = Find-Module -Name $module.name -Repository PSGallery -ErrorAction Stop
         }
         Catch {
            Write-Warning "Module $($module.name) was not found in the PSGallery"
         }
    
         #compare versions
         if ($online.version -gt $module.version) {
            $UpdateAvailable = $True
         }
         else {
            $UpdateAvailable = $False
         }
    
         #write a custom object to the pipeline
         [pscustomobject]@{
            Name = $module.name
            MultipleVersions = ($g.name -contains $module.name)
            InstalledVersion = $module.version
            OnlineVersion = $online.version
            Update = $UpdateAvailable
            Path = $module.modulebase
         }
     
    } #foreach

    Write-Host "Check complete" -ForegroundColor Green
}

function Set-ChromeDownloadFolder ($path) {
    # https://stackoverflow.com/questions/53505079/how-to-change-default-download-folder-in-chrome-using-powershell
    $UserFolder = "$env:USERPROFILE"
    if ($UserFolder -like "*LADM") { $UserFolder = $UserFolder.Split("LADM")[0] }
    # Test if Chrome is open (must be closed for this to work)
    # Get Chrome Preferences.json and alter the "download.default_directory" value
    $File = "$UserFolder\AppData\Local\Google\Chrome\User Data\Default\Preferences"
    $Chrome = Get-Content $File | ConvertFrom-Json
    $Chrome.download.default_directory = $path   # "C:\0\Downloads"    # "$UserFolder\Downloads"
    $Chrome | ConvertTo-Json | Out-File -FilePath $File -Force
}

function Uninstall-ModuleIfNotLatest ($ModuleName) {
    $Latest = Get-InstalledModule $ModuleName   # Just returns the latest version only
    Test-ModuleUpdate $ModuleName

    "`nIf proceed with uninstall of older versions, the following will run:"
    "Get-InstalledModule $ModuleName -AllVersions | ? { $_.Version -ne $Latest.Version } | Uninstall-Module"
    Get-InstalledModule $ModuleName -AllVersions | ? { $_.Version -ne $Latest.Version } | Uninstall-Module -WhatIf

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes";
    $no  = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No";
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no);
    $caption = ""   # Did not need this before, but now getting odd errors without it.
    $answer = $host.ui.PromptForChoice($caption, "Proceed with Module uninstall?", $choices, 1)   # Set to 0 to default to "yes" and 1 to default to "no"

    if ($answer -eq 0) {
        Get-InstalledModule $ModuleName -AllVersions | ? { $_.Version -ne $Latest.Version } | Uninstall-Module
    }
    else {
        "No changes made."
    }
}

function Help-OneDrive {
    Write-Host ''
    Write-Host 'Disable OneDrive and remove its icon from File Explorer, or restore, with these'
    Write-Host 'Type GPedit.msc and hit Enter or OK to open Local Group Policy Editor.'
    Write-Host 'Local Computer Policy -> Computer Configuration -> Administrative Templates -> Windows Components -> OneDrive.'
    Write-Host 'In the right pane, double click on policy named Prevent the usage of OneDrive for file storage.'
    Write-Host 'https://techjourney.net/disable-or-uninstall-onedrive-completely-in-windows-10/'
    Write-Host ''
    Write-Host 'Convert the below to PowerShell commands and test removal and then restore'
    Write-Host ''
    Write-Host 'Windows Registry Editor Version 5.00'
    Write-Host '; 64-bit Hide OneDrive From File Explorer'
    Write-Host '[HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]'
    Write-Host '"System.IsPinnedToNameSpaceTree"=dword:00000000'
    Write-Host '[HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]'
    Write-Host '"System.IsPinnedToNameSpaceTree"=dword:00000000'
    Write-Host ''
    Write-Host 'Windows Registry Editor Version 5.00'
    Write-Host '; 64-bit Restore OneDrive to File Explorer'
    Write-Host '[HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]'
    Write-Host '"System.IsPinnedToNameSpaceTree"=dword:00000001'
    Write-Host '[HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]'
    Write-Host '"System.IsPinnedToNameSpaceTree"=dword:00000001'
    Write-Host ''
    Write-Host 'Windows Registry Editor Version 5.00'
    Write-Host '; 32-bit Hide OneDrive From File Explorer'
    Write-Host '[HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]'
    Write-Host '"System.IsPinnedToNameSpaceTree"=dword:00000000'
    Write-Host ''
    Write-Host 'Windows Registry Editor Version 5.00'
    Write-Host '; 32-bit Restore OneDrive to File Explorer'
    Write-Host '[HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]'
    Write-Host '"System.IsPinnedToNameSpaceTree"=dword:00000001'
    Write-Host ''
}



function Help-ToolkitConfig {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host 'Toolkit Summary' -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host ""
    Write-Host "To update to the latest version of 'ProfileExtensions.ps1' and 'Custom-Tools.psm1':" -ForegroundColor Green
    Write-Host "   . Update-Toolkit" -ForegroundColor Yellow
    Write-Host "This function simply redirects to:"
    Write-Host "   iex ((New-Object System.Net.WebClient).DownloadString('https://git.io/JqCtf'))"
    Write-Host ""
    Write-Host "ProfileExtensions is placed in the `$profile folder and is called on all new console sessions." -ForegroundColor Green
    Write-Host "Type 'more `"`$(`$Profile)_extensions.ps1`"' to review all contents."
    Write-Host ""
    Write-Host "To view the profile extensions:   " -ForegroundColor Green -NoNewline ; Write-Host "cd (Split-Path `$Profile); dir" -ForegroundColor Yellow
    Write-Host "The profile extensions includes various essential aliases and core functions (~0.1 sec load time)." -ForegroundColor Green
    Write-Host "To view functions from 'ProfileExtensions.ps1':   " -ForegroundColor Green -NoNewline ; Write-Host "myfunctions" -ForegroundColor Yellow
    Write-Host "To view functions from 'Custom-Tools.psm1'    :   " -ForegroundColor Green -NoNewline ; Write-Host "mod Custom-Tools" -ForegroundColor Yellow
    Write-Host "To elevate (sudo) a function: run 'sudo' to elevate the current shell, or 'sudo <command> to elevate any given command supplied." -ForegroundColor Green
    Write-Host ""
    Write-Host "Help Extensions:   m <string>   # advanced man/help that includes all about_* documentation, and <Alias|Cmdlet|Function|ExternalScript|Application> info" -ForegroundColor Green
    Write-Host "   e.g.  m Para      # Try exactly this to show all variants of 'Para' in help files, note no need for wildcards."
    Write-Host "  'm' is a help wrapper tool that has much more functionality. Type 'm' on its own for more information."
    Write-Host "def <command>   # show the definitions for any command: Alias, Cmdlet, Function, ExternalScript, Application"
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Green
    Write-Host "   mods        # All currently installed modules and installed locations."
    Write-Host "   mods W*     # Show modules starting with 'w' and their installed locations."
    Write-Host "   mod <name>  # See details on a specific module"
    Write-Host "   mod <name> <function_in_mod>   # Show the 'def' (definitions) for function_in_mod that is in the module."
    Write-Host " e.g.  mod Custom-Tools"
    Write-Host "       mod Custom-Tools Help-sls"
    Write-Host "   def sudo    # To view sudo definition"
    Write-Host ""
    Write-Host "'cd' includes additional functionality over Set-Location" -ForegroundColor Green
    Write-Host "Type 'cd' on its own to see additional Set-Location functionality."
    Write-Host
    Write-Host "More Examples (type 'myfunctions' or 'mod custom-tools' for more infor, or 'def uptime' etc):" -ForegroundColor Green
    Write-Host "cd.. is a function in PowerShell (test that by using 'def cd..'), so have defined cd... and cd.... in same way."
    Write-Host "uptime (Show uptime of computer), get-errors (get count of recent Event Log errors), touch `$file"
    Write-Host "tls, tls12 (set TLS, important for GitHub), securityprotocol to view status"
    Write-Host "DateTimeNowString (get datetime as filename compatible string), DateTimeNowStringMinutes (if don't need seconds)"
    Write-Host "BackupProfile (use robocopy to backup current Profile folder to D:\Backup\`$env:USERNAME_`$DateTimeNowStringMinutes"
    Write-Host "sys function (get summary of most used system details)"
    Write-Host "Various 'prompt' functions are included in Custom-Tools. Type 'prompt' then his Ctrl-Space to view." -ForegroundColor Green
    # Write-Host "Use 'more `"`$(`$Profile)_extensions.ps1'`" to review all contents." -ForegroundColor Green
    Write-Host ""
    Write-Host "Custom-Tools.psm1 Module installed to user Module folder." -ForegroundColor Green
    Write-Host "Keep this mainly for usually longer functions that will only be fully loaded when the Module is called." -ForegroundColor Green
    Write-Host "e.g. dirsize function (uses robocopy to calculate folder size, which is about 6x faster than native PowerShell)"
    Write-Host "     dirsizeps function (same as dirsize but using native PowerShell)"
    Write-Host "     Get-DiskFree function (Unix df equivalent, check all volumes for disk used and free)"
    Write-Host ""
    Write-Host "Chocolatey:" -ForegroundColor Green
    Write-Host "This is installed by default as a PowerShell tool very important for package management tasks."
    Write-Host "choco list -lo             # View locally installed packages."
    Write-Host "choco search <string>      # Find chocolatey packages."
    Write-Host "choco info <packagename>   # To get detailed info on a chocolatey package."
    Write-Host "choco inst <packagename>   # To install a chocolatey package."
    Write-Host ""
    # Write-Host "Scoop is not installed by default, use Install-Sccop from Cutsom-Tools" -ForegroundColor Green
    # Write-Host "scoop search / info / install <packagename>   # search, info, install"
    # Write-Host ""
    # Write-Host "Review script actions above if required as this window will close after this" -ForegroundColor White -BackgroundColor Red
    # Write-Host "point if this script was called from another console." -ForegroundColor White -BackgroundColor Red
    Write-Host ""
    Write-Host "Some quick commands to try:" -ForegroundColor Green
    Write-Host "   mods        # All currently installed modules and installed locations."
    Write-Host "   mods W*     # Show modules starting with 'w' and their installed locations."
    Write-Host "   mod <name>  # See details on a specific module"
    Write-Host "   mod <name> <function_in_mod>   # Show the 'def' (definitions) for function_in_mod that is in the module."
    Write-Host " e.g.  mod Custom-Tools"
    Write-Host "       mod Custom-Tools Help-sls"
    Write-Host ""
    Write-Host "Help Tools" -ForegroundColor Green
    Write-Host "   def <command>   # show the definitions for any command: Alias, Cmdlet, Function, ExternalScript, Application"
    Write-Host "e.g.  def Install-NotepadPlusPlus"
    Write-Host "   m <Alias|Cmdlet|Function|ExternalScript|Application>"
    Write-Host "e.g.  m Para      # Try exactly this to show all variants of 'Para' in help files, note no need for wildcards."
    Write-Host "  'm' is a help wrapper tool that has much more functionality. Type 'm' on its own for more information."
    Write-Host ""
    try { & "$Home\Desktop\MySandbox\MyPrograms\winfetch.ps1" } catch { "WinFetch hit an error and could not load.`n" }
    # Write-Host "Final prompt in case this was called in a separate console (which will immediately close"
    # Write-Host "after you continue - in which case, make sure to check the above output if required)."
    # Write-Host "Press Enter to continue...:" ; cmd /c pause | out-null   # PowerShell v2 compatible version of 'Pause'

    # Good general overview of some basic PowerShell
    # https://mathieubuisson.github.io/powershell-linux-bash/

    ### choco inst snappy-driver-installer   # Portable, will test thousands of drivers against an install (about 19 GB if download all components)
    ### choco install cascadiacode (maybe not)
    ### choco install ethanbrown.conemuconfig

    # Some Windows Tricks (from Edwin):
    # ipconfig | clip   # Puts output in your clipboard
    # In Explorer, type cmd or powershell into the address bar to open a console at the current location
    # Regedit -m   # Multiple regedit open
    # A funny message: helpmsg 4006
    # When cmd.exe is disabled by a policy open task manager and then ctrl cmd opens a cmd.exe
    # Without ctrl the run doesn't open it
}
Set-Alias Help-CustomTools Help-ToolkitConfig

function Help-PowershellConsole {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host 'Console Settings Tips / Notes' -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    $out  = "Ideally, I only want to always open an Admin console shortcut in my TaskBar and Start Menu. "
    $out += "This means that by deafult, this means always seeing a UAC (User Account Control) prompt. "
    $out += "There are ways to prevent the UAC but they can be complex (and overall it is better not to "
    $out += "not disable UAC as it can expose the system to various malware/ransomware)."
    Write-Wrap $out ; Write-Host ""
    Write-Host "The method I use for this is simple and is done in a few seconds: "
    Write-Host "- Start menu > right-click on the PowerShell shortcut > More > Open file location"
    Write-Host "- Make a copy of the default PowerShell shortcut and call it 'PowerShell (Admin)'"
    Write-Host "- Right-click on 'PowerShell (Admin) > Properties > Shortcut tab > Advanced"
    Write-Host "- Check the 'Run as administrator' option in here and then press OK and OK again to close Properties"
    Write-Host "- Right-click on the shortcut two more times, once to 'Pin to Start' and once to 'Pin to taskbar'"
    Write-Host "Note Alt+y, Alt+j for the UAC prompt"
    Write-Host ""
    Write-Host "To achieve the much more advanced goal of always opening a PowerShell Console as Admin"
    Write-Host "*without* seeing a UAC prompt and *without* disabling UAC itself, follow the notes here:"
    Write-Host "https://www.winhelponline.com/blog/run-programs-elevated-without-getting-the-uac-prompt/"
    Write-Host "https://winaero.com/blog/open-any-program-as-administrator-without-uac-prompt/"
    Write-Host "https://www.youtube.com/watch?v=Jg2UFoMWDB4"
    Write-Host "https://invoke-thebrain.com/2019/04/runas-without-runas-exe/"
    Write-Host "https://gallery.technet.microsoft.com/scriptcenter/How-to-easily-run-an-0c0eb47a"
    Write-Host "http://woshub.com/run-program-without-admin-password-and-bypass-uac-prompt/"
    Write-Host ""
    Write-Host "ToDo: possibly create a function that will automate all of the above steps."
    Write-Host "https://forums.hak5.org/topic/45439-powershell-real-uac-bypass/"
    Write-Host "https://raw.githubusercontent.com/PoSHMagiC0de/Invoke-TaskCleanerBypass/master/Invoke-TaskCleanerBypass.ps1"
    Write-Host "It uses dynamic parameters and can take in the standard posh base64 encoded commands or a file location of your script."
    Write-Host "Create encoded stager to downloadstring the bypass script from web server and execute with "Invoke-Expression" IEX for short with the command."
    Write-Host "You probably can take this function, add after it the command to run it with your parameters and encode the whole thing to run.  No bypass to execution policy needed.  Anyway, look at the script.  Some modifications were needed to the reg hack.  I needed to use cmd /c in front so I could escape the appended stuff that gets added when ran like the cleaner command."
    Write-Host "That was breaking the exploit.  So the new reg entry is:"
    Write-Host "    cmd /c yourpayload & ::"
    Write-Host "That runs the command and then rems out whatever else is there. SQL injection for registries."
    Write-Host "Since I won the competition this month so I am not payloading this. Someone else can run with this and create a BB payload."
    Write-Host "I know a few ways to use it but someone else can have a turn."
    Write-Host "FYI: It checks if you have Win10, member of local admins and already UAC bypassed."
    Write-Host "Will run if bypassed, will do nothing if not on 10 or greater and/or not a local admin."
    Write-Host ""
    # Quicksudo ...
    # function sudo {
    #     $command = "powershell -noexit " + $args + ";#";
    #     
    #     Set-ItemProperty -Path "HKCU:\Environment" -Name "windir" -Value $command ;
    #     schtasks /run /tn \Microsoft\Windows\DiskCleanup\SilentCleanup /I;
    #     Remove-ItemProperty -Path "HKCU:\Environment" -Name "windir"
    # }
    #
    # Disable UAC (don't do this, bad idea)
    # New-ItemProperty -Path HKLM:Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -PropertyType DWord -Value 0 -Force
    # Restart-Computer
}



function Help-AsciiArt {

@'

https://stackoverflow.com/questions/35022078/how-do-i-output-ascii-art-to-console

# Only works on PS 5.1, not 7.2.4
$Host.UI.RawUI.WindowTitle = "Windows Powershell " + $Host.Version;

# Emoji's only display in a Terminal that supports them like Windows Terminal
Invoke-RestMethod -Uri http://wttr.in/Amsterdam?format=2 -UseBasicParsing -DisableKeepAlive

# Very good, implement this! The Lonely Administrator
https://jdhitsolutions.com/blog/powershell/8163/friday-fun-powershell-weather-widget/
https://jdhitsolutions.com/blog/powershell-tips-tricks-and-advice/
https://gist.github.com/jdhitsolutions/f2fb0184c2dbab107f2416fb775d462b

# Also very good, use this!
http://vcloud-lab.com/entries/powershell/powershell-trick-convert-text-to-ascii-art
https://github.com/kunaludapi/Powershell-trick-Convert-text-to-ASCII-Art
http://vcloud-lab.com/files/documents/Convertto-TextASCIIArt.ps1

# Write-ASCII function
https://www.powershelladmin.com/wiki/Ascii_art_characters_powershell_script.php

Write-Host -ForegroundColor DarkYellow "                       _oo0oo_"
Write-Host -ForegroundColor DarkYellow "                      o8888888o"
Write-Host -ForegroundColor DarkYellow "                      88`" . `"88"
Write-Host -ForegroundColor DarkYellow "                      (| -_- |)"
Write-Host -ForegroundColor DarkYellow "                      0\  =  /0"
Write-Host -ForegroundColor DarkYellow "                    ___/`----'\___"
Write-Host -ForegroundColor DarkYellow "                  .' \\|     |// '."
Write-Host -ForegroundColor DarkYellow "                 / \\|||  :  |||// \"
Write-Host -ForegroundColor DarkYellow "                / _||||| -:- |||||- \"
Write-Host -ForegroundColor DarkYellow "               |   | \\\  -  /// |   |"
Write-Host -ForegroundColor DarkYellow "               | \_|  ''\---/''  |_/ |"
Write-Host -ForegroundColor DarkYellow "               \  .-\__  '-'  ___/-. /"
Write-Host -ForegroundColor DarkYellow "             ___'. .'  /--.--\  `. .'___"
Write-Host -ForegroundColor DarkYellow "          .`"`" '<  `.___\_<|>_/___.' >' `"`"."
Write-Host -ForegroundColor DarkYellow "         | | :  `- \`.;`\ _ /`;.`/ - ` : | |"
Write-Host -ForegroundColor DarkYellow "         \  \ `_.   \_ __\ /__ _/   .-` /  /"
Write-Host -ForegroundColor DarkYellow "     =====`-.____`.___ \_____/___.-`___.-'====="
Write-Host -ForegroundColor DarkYellow "                       `=---='"

function Get-Funky{
    param([string]$Text)

    # Use a random colour for each character
    $Text.ToCharArray() | ForEach-Object{
        switch -Regex ($_){
            # Ignore new line characters
            "`r"{
                break
            }
            # Start a new line
            "`n"{
                Write-Host " ";break
            }
            # Use random colours for displaying this non-space character
            "[^ ]"{
                # Splat the colours to write-host
                $writeHostOptions = @{
                    ForegroundColor = ([system.enum]::GetValues([system.consolecolor])) | get-random
                    # BackgroundColor = ([system.enum]::GetValues([system.consolecolor])) | get-random
                    NoNewLine = $true
                }
                Write-Host $_ @writeHostOptions
                break
            }
            " "{Write-Host " " -NoNewline}

        } 
    }
}

$art = " .:::.   .:::.`n:::::::.:::::::`n:::::::::::::::
':::::::::::::'`n  ':::::::::'`n    ':::::'`n      ':'"
Get-Funky $art 

# https://jdhitsolutions.com/blog/powershell/7278/friday-fun-powershell-ascii-art/
$t = @"
  _____                       _____ _          _ _ 
 |  __ \                     / ____| |        | | |
 | |__) |____      _____ _ __ (___ | |__   ___| | |
 |  ___/ _ \ \ /\ / / _ \ '__\___ \| '_ \ / _ \ | |
 | |  | (_) \ V  V /  __/ |  ____) | | | |  __/ | |
 |_|   \___/ \_/\_/ \___|_| |_____/|_| |_|\___|_|_|
                                                   
"@

for ($i=0;$i -lt $t.length;$i++) {
    if ($i%2) {
        $c = "red"
    }
    elseif ($i%5) {
        $c = "yellow"
    }
    elseif ($i%7) {
        $c = "green"
    }
    else {
        $c = "white"
    }
    write-host $t[$i] -NoNewline -ForegroundColor $c
}

# Also note the Figlet (and dependency Pansies) modules
Uses standard figlet fonts http://www.figlet.org/examples.html
Write-Figlet [-Message] <String> [[-Font] <String>] [[-LayoutRule] {FullSize | Fitting | Smushing | Custom}] [[-ColorChars] <String>] [[-Foreground] <RgbColor[]>]
    [[-Background] <RgbColor[]>] [[-Colorspace] <String>] [<CommonParameters>]
def Pansies
Get-ColorWheel, Get-Complement, Get-Gradient, New-Hyperlink, New-Text, Write-Host

'@ | more
}

function Help-WindowsTools {
# function cpl / sysdm / control / rundll
# Create a function that opens all of these "sysdm" on it's own shows options, then "sysdm env" would open:
# maybe setup some kind of 'sys' function for all of these weird and wonderful items.
# sys add/remove/uninstall/programs => appwiz.cpl
# sys => System Properties + Control Panel + System Settings + Environment Variables
# sys computer => Computer Name tab
# sys local/security/policy/localsecuritypolicy
# sys passwords/user => netplwiz.exe
# sys hardware => Hardware tab + Device Manager + Devices and Printers etc etc
# Also should sort of say what command is being run *just write it below
# Maybe just a bunch of switches? sys -systemproperties
# # 
@"

List of command line for various Windows Tools:
secpol.msc   # Local Security Policy
appwiz.cpl   # Programs and Features 
netplwiz.exe # User Accounts (can set unattended logon password here)

# https://www.tenforums.com/tutorials/77458-rundll32-commands-list-windows-10-a.html
rundll32 sysdm.cpl,EditEnvironmentVariables   # Open Environment Variables dialogue
Maybe just similar to cc / go with a hash table

 Control panel tool                                     Command
-----------------------------------------------------------------
rundll32.exe shell32.dll,Control_RunDLL                # Open Control Panel (old form)
rundll32.exe shell32.dll,Control_RunDLL srchadmin.dll  # Indexing Options
rundll32 sysdm.cpl,EditEnvironmentVariables            # Open Environment Variables dialogue
RunDll32.exe InetCpl.cpl,ClearMyTracksByProces         # Clear all browsing history (dangerous, IE only? Test this)

SystemPropertiesComputerName # Computer Name tab
    cmd.exe /c "sysdm.cpl,1"   # or just sysdm.cpl,1 in cmd.exe
SystemPropertiesHardware     # Hardware tab
    cmd.exe /c "sysdm.cpl,2"
SystemPropertiesAdvanced     # Advanced tab
    cmd.exe /c "sysdm.cpl,3"
    SystemPropertiesPerformance  # Performance Options (dialogue under Avvanced tab)
    SystemPropertiesDataExecutionPrevention # Data Execution Prevention tab on Performance Options
SystemPropertiesProtection   # Protection tab 
    cmd.exe /c "sysdm.cpl,4"
SystemPropertiesRemote       # Remote tab
    cmd.exe /c "sysdm.cpl,5"

sysdm              # System Properties (control sysdm.cpl)
appwiz             # Add/Remove Programs (control appwiz.cpl)
access.cpl         # Accessibility Options (control access.cpl)
timedate           # Date and Time (control timedate.cpl)
desk               # Display Settings (control desk.cpl)
control fonts      # Fonts Folder (note: no ".cpl" and requires "control")
intl               # Region (control intl.cpl)

devmgmt            # Device Manager (devmgmt.msc)
main               # Mouse Properties (control main.cpl)
main keyboard      # Keyboard Properties (control main.cpl keyboard)
msys               # Multimedia Properties (control mmsys.cpl)
mmsys sounds       # Sound Properties (control mmsys.cpl sounds)
control printers   # Devices and Printers
control sticpl.cpl # Scanners and Cameras (only this exact syntax works)
joy                # Joystick Properties (control joy.cpl)
inetcpl            # Internet Properties (control inetcpl.cpl)
ncpa               # Network Adapter Settings (also: control netconnections)
Good article on netsh commands
https://www.windowscentral.com/how-manage-wireless-networks-using-command-prompt-windows-10

powercfg.cpl       # Power Management (control powercfg.cpl)
powercfg.exe       # Command-line Power Management tool

[ mlcfg32         # Microsoft Exchange (Windows Messaging) (removed in newer Win 10) ]
[ wgpocpl         # Microsoft Mail Post Office (removed in newer Win 10) ]
[ modem           # Modem Properties (removed from newer Win 10) ]
[ main pc card    # PC Card / PCMCIA (removed from newer Win 10) ]
[ findfast        # FindFast (control findfast.cpl) ]

# ; Run, control.exe sysdm.cpl,,1 ; Computer Name, SystemPropertiesComputerName.exe
# ; Run, control.exe sysdm.cpl,,2 ; Hardware, SystemPropertiesHardware.exe
# ; Run, control.exe sysdm.cpl,,3 ; Advanced, SystemPropertiesAdvanced.exe
# ; Run, control.exe sysdm.cpl,,4 ; System Protection, SystemPropertiesProtection.exe
# ; Run, control.exe sysdm.cpl,,5 ; Remote, SystemPropertiesRemote.exe
# Run, control password.cpl ; Password Properties
# Run, control intl.cpl ; Regional Settings

# Run, control inetcpl.cpl ; Internet Properties
# Run, control netcpl.cpl ; Network Properties
# Run, control netconnections
# return
# 
# :*:cccd::   ; Devices, Mouse, Keyboard, Joystick, Printers, etc
# Run, control hdwwiz.cpl   ; Device Manager
# Run, control desk.cpl ; Display Properties
# Run, control powercfg.cpl ; Power Management (Windows 98)
# Run, control printers ; Printers Folder
# Run, control mmsys.cpl ; Multimedia Properties
# Run, control main.cpl keyboard ; Keyboard Properties
# Run, control main.cpl ; Mouse Properties
# Run, control joy.cpl ; Joystick Properties
# Run, control mmsys.cpl sounds ; Sound Properties
# Run, control sticpl.cpl ; Scanners and Cameras
"@ | more
}

function Help-Sandbox {
    $out = @'

# Windows Sandbox
##########

*** 1. Confirm Virtualization UEFI/BIOS for the motherboard and Task Manager Performance tab, lower right, "Virtualization: Enabled".

*** 2. Enable Windows Sandbox from an Admin PowerShell console:
    Enable-WindowsOptionalFeature -FeatureName "Containers-DisposableClientVM" -All -Online -NoRestart   # Enable Windows Sandbox

*** 3. You must reboot to complete installation, the -NoRestart above is just to suppress the reboot question at the end.

The Sandbox will always open only a pristine Windows 10 environment with user "WDAGUtilityAccount".
The Sandbox is always wiped on closing and uses Windows Container technology so that the VM only takes up
~25MB closed and 100MB when running. It appears as a *non-activated* version of Windows 10 that has the sam
Build veriosn as the host OS. Note that the Sandbox cannot co-exist with VirtualBox or VMware. You should try
to only use Hyper-V on a system using the Sandbox.
https://techcommunity.microsoft.com/t5/windows-kernel-internals/windows-sandbox-config-files/ba-p/354902?WT.mc_id=thomasmaurer-blog-thmaure
https://www.ghacks.net/2019/04/26/install-the-windows-sandbox-in-windows-10-home/
https://techcommunity.microsoft.com/t5/windows-kernel-internals/windows-sandbox/ba-p/301849/page/2
https://docs.microsoft.com/en-us/virtualization/windowscontainers/about/

Dynamically generated Image
At its core Windows Sandbox is a lightweight VM, but Sandbox key enhancement is ability to use the Windows 10 installed on your computer,
instead of downloading a new VHD image as you would have to do with an ordinary virtual machine.
The challenge is that some OS files can change. The solution is to construct what we refer to as "dynamic base image": an OS image that
has clean copies of files that can change, but links to files that cannot change that are in the Windows image that already exists on the
host. The majority of the files are links (immutable files) and that's why the small size (~100MB) for a full operating system. We call
this instance the "base image" for Windows Sandbox, using Windows Container parlance.
When Sandbox is not in use, the dynamic base image is compressed (~25MB). When installed the dynamic base package is only 100MB disk space.

*** Main uses:
- Test software that you are unsure about.
- Open a web page that you are unsure about.

*** Start a LogonCommand using .wsb Configuration file with MappedFolder and <LogonCommand> settings
LogonCommand runs a DOS commands, so want to enable this to run a PowerShell command.
For this example, create 3 files in C:\Sandbox, Sandbox.wsb, Sandbox.cmd, Sandbox.ps1

*** "C:\Sandbox\Sandbox.cmd"
powershell.exe -noexit -executionpolicy remotesigned -file "C:\Users\WDAGUtilityAccount\Desktop\Sandbox\Sandbox.ps1"
:: https://stackoverflow.com/questions/59546857/calling-a-powershell-script-from-batch-script-but-need-the-batch-script-to-wait
:: https://www.howtogeek.com/204088/how-to-use-a-batch-file-to-make-powershell-scripts-easier-to-run/
:: Note that the Sandbox has the default execution policy to block all scripts. Without the execution policy switch
:: there will be no output or errors, the script just fails.

*** "C:\Sandbox\Sandbox.wsb" (also adding my OneDrive as a share, note that all shares go to the Sandbox Desktop):
<Configuration>
<MappedFolders>
  <MappedFolder>
    <HostFolder>C:\Sandbox</HostFolder>
    <ReadOnly>false</ReadOnly>
  </MappedFolder>
  <MappedFolder>
    <HostFolder>D:\0 Cloud\OneDrive</HostFolder>
    <ReadOnly>false</ReadOnly>
  </MappedFolder>
</MappedFolders>
<LogonCommand>
  <Command>C:\Users\WDAGUtilityAccount\Desktop\Sandbox\Sandbox.cmd</Command>
</LogonCommand>
</Configuration>

*** "C:\Sandbox\Sandbox.ps1" put anything required in here, e.g.
del "C:\Users\WDAGUtilityAccount\Desktop\Microsoft Edge.lnk"   # Delete Edge from Desktop
$vscode_url  = "https://update.code.visualstudio.com/latest/win32-x64-user/stable"
$vscode_file = "C:\Users\WDAGUtilityAccount\Desktop\Sandbox\vscode.exe"
if (!(Test-Path $vscode_file)) { curl.exe -L $vscode_url --output $vscode_file }   # Download latest vscode.exe if it does not already exist
# Install and run VSCode:
& $vscode_file /verysilent /suppressmsgboxes


To run the above configuration, double-click on Sandbox.wsb to start the Sandbox with the above configuration.

To disable, or to install from DOS:
    Disable-WindowsOptionalFeature -FeatureName "Containers-DisposableClientVM" -Online -NoRestart # Disable Windows Sandbox
        Dism /online /Enable-Feature /FeatureName:"Containers-DisposableClientVM" -All             # Enable in DOS
        Dism /online /Disable-Feature /FeatureName:"Containers-DisposableClientVM"                 # Disable in DOS
'@
echo $out | more
}

function Help-HyperV {
    $out = @'

# Hyper-V
##########

*** 1. Enable Hyper-V from an Admin PowerShell console:
    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All -NoRestart
Note: You must reboot before final installation, the -NoRestart is just to suppress the reboot question at the end.
After reboot, you can examine the host:
    Get-VMhost | Select *

*** 2. Create and start a VM:
    New-VM -Name MyVM -path D:\VMs\MyVM_Files --MemoryStartupBytes 2048MB
Create a new virtual hard disk:
    New-VHD -Path D:\VMs\MyVM\MyVM.vhdx -SizeBytes 10GB -Dynamic
Attach our new virtual hard disk to the VM:
    Add-VMHardDiskDrive -VMName MyVM -path "D:\VMs\MyVM_Files\MyVM\MyVM.vhdx"
Map an ISO image to VM CD/DVD (which will be used to install the OS inside the VM, so this can be a Windows or Linux DVD image):
    Set-VMDvdDrive -VMName -ControllerNumber 1 -Path
Start the new VM:
    Start-VM -Name
    Start-VM -Name <virtual machine name>     # Start a VM
    Get-VM | where {$_.State -eq 'Running'}   # All running VMs
    Get-VM | where {$_.State -eq 'Off'}       # All powered off VMs
    Get-VM | where {$_.State -eq 'Off'} | Start-VM     # Start all currently powered off VMs
    Get-VM | where {$_.State -eq 'Running'} | Stop-VM  # Stop all running VMs
    Get-VM -Name <VM Name> | Checkpoint-VM -SnapshotName <name for snapshot>   # Create a VM checkpoint
An alternative way to 
    $VMName = "VMNAME"
    $VM = @{
        Name = $VMName
        MemoryStartupBytes = 2147483648
        Generation = 2
        NewVHDPath = "C:\Virtual Machines\$VMName\$VMName.vhdx"
        NewVHDSizeBytes = 53687091200
        BootDevice = "VHD"
        Path = "C:\Virtual Machines\$VMName"
        SwitchName = (Get-VMSwitch).Name
    }
    New-VM @VM

*** 3. Import a VM from a previous setup:
I had an existing VM in D:\VMs\Win7 (I keep all of my VMs there).
The .vmcx is the file that has to be imported (under the "Virtual Machines" folder).
    Import-VM "D:\VMs\Win7\Virtual Machines\62ECC257-F4F8-4D94-8ECE-D9135EB695D6.vmcx"

    Name State CPUUsage(%) MemoryAssigned(M) Uptime   Status             Version
    ---- ----- ----------- ----------------- ------   ------             -------
    Win7 Saved 0           0                 00:00:00 Operating normally 9.0

Can also use:
    Import-VM "D:\VMs\Win7\Virtual Machines\62ECC257-F4F8-4D94-8ECE-D9135EB695D6.vmcx" -Copy -GenerateNewId
-Copy will copy all files into the default Hyper-V folder.
-GenerateNewId is useful when importing multiple VMs to avoid ID conflicts
    
Example of a bulk import:
    $VMlist = Get-ChildItem D:\VM-Exports -recurse -include *.exp
    $VMlist | foreach { Import-VM -path $_.Fullname -Copy -VhdDestinationPath $VMDefaultDrive -VirtualMachinePath $VMDefaultPath -SnapshotFilePath $VMDefaultPath -SmartPagingFilePath $VMDefaultPath -GenerateNewId }
When done, simply attach these machines to the current network using the Connect-VMSwitch cmdlet.

Default locations for Virtual Machines on installing Hyper-V are:
    Get-VMHost | select VirtualHardDiskPath, VirtualMachinePath | fl
    
    VirtualHardDiskPath : C:\Users\Public\Documents\Hyper-V\Virtual Hard Disks
    VirtualMachinePath  : C:\ProgramData\Microsoft\Windows\Hyper-V

Note on Generation 1 vs Generation 2. I always use Generation 1.
Generation 2 has few very important improvements and is a nightmare to boot from USB sticks or ISO's (as it is UEFI/GPT only).
    UEFI Firmware, New virtual hardware for network adapters, New virtual hardware for input devices (mouse, PS/2 & i8042 keyboards)
    New virtual hardware for video adapter, IDE controller replaced by SCSI controller, PCI Bus has been removed, Floppy drive support has been removed
    New support for booting from SCSI drive or Network Adapter, GPT disk partitions supported, RemoteFX is no longer supported
    Get-VM | Select Name,Generation
https://docs.microsoft.com/en-gb/archive/blogs/ausoemteam/deciding-when-to-use-generation-1-or-generation-2-virtual-machines-with-hyper-v

https://www.tenforums.com/tutorials/2087-hyper-v-virtualization-Install-use-windows-10-a.html
https://www.sysnettechsolutions.com/en/create-new-virtual-machine-in-hyper-v-manager/
https://www.altaro.com/hyper-v/hyper-v-attach-existing-virtual-disk/
Add-VMHardDiskDrive -VMName svtest -ControllerType IDE -ControllerNumber 1 -ControllerLocation 1 -Path 'C:\LocalVMs\Virtual Hard Disks\disk2.vhdx'

Get-WindowsOptionalFeature -Online -FeatureName *hyper-v* | select DisplayName, FeatureName
Shows a flat display of the hierarchical tree that you get if opening the Windows Features window.
DisplayName                           FeatureName
-----------                           -----------
Hyper-V                               Microsoft-Hyper-V-All
Hyper-V Platform                      Microsoft-Hyper-V
Hyper-V Management Tools              Microsoft-Hyper-V-Tools-All
Hyper-V Module for Windows PowerShell Microsoft-Hyper-V-Management-PowerShell
Hyper-V Hypervisor                    Microsoft-Hyper-V-Hypervisor
Hyper-V Services                      Microsoft-Hyper-V-Services
Hyper-V GUI Management Tools          Microsoft-Hyper-V-Management-Clients

# Install only the PowerShell module
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell

# Install the Hyper-V management tool pack (Hyper-V Manager and the Hyper-V PowerShell module)
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Tools-All

# Install the entire Hyper-V stack (hypervisor, services, and tools)
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All

https://www.red-gate.com/simple-talk/sysadmin/powershell/hyper-v-powershell-basics/
https://www.red-gate.com/simple-talk/sysadmin/powershell/hyper-v-and-powershell-shielded-virtual-machines/
'@
echo $out | more
}

function Help-WSL {

$out = @'

# WSL (Windows Subsystem for Linux) (Requires Reboot):
##########

*** 1. Enable the Windows Optional Feature for WSL from an Admin PowerShell console:
    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -NoRestart
Note: You must reboot before final installation, the -NoRestart is just to suppress the reboot question at the end.

*** How to get onto WSL 2 ***. by default, if not on Fast Ring, WSL 1 will be installed (as of current 2020-03 1909 build).
https://scotch.io/bar-talk/trying-the-new-wsl-2-its-fast-windows-subsystem-for-linux
This will not be required when WSL is in the main build.
Requirements for installing WSL 2:
- Hyper-V capable computer.
- WSL installed already (as above).
- Windows 10 version 18917 or greater (currently this is Fast Ring only, running ver from DOS showed Version 10.0.18363.720).
  To get on the Fast Ring, Settings > Update > Windows Insider Program (can also search for "Windows Insider Program" from Start menu).
  Note that this requires a Hotmail account linked to this system (which I don't usually do).
Once the above are in place, install WSL 2 "VirtualMachinePlatform"
    Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform
    wsl --set-default-version 2


*** 2. Download a Distro
To download distros using PowerShell, use the Invoke-WebRequest cmdlet. Here's a sample instruction to download Ubuntu 16.04.
Invoke-WebRequest -Uri https://aka.ms/wsl-ubuntu-1604 -OutFile Ubuntu.appx -UseBasicParsing
Tip: If the download is taking a long time, turn off the progress bar by setting $ProgressPreference = 'SilentlyContinue'
Download using curl (Win 10 Spring 2018 update or later includes a full curl.exe version to invoked web requests (i.e. HTTP GET, POST, PUT, etc. commands)
Note: this example must use 'curl.exe' to ensure that, in PowerShell, it uses the executable and not the PowerShell curl alias (for Invoke-WebRequest)
    curl.exe -L -o ubuntu-1804.appx https://aka.ms/wsl-ubuntu-1804
    curl.exe -L -o ubuntu-1804.appx https://aka.ms/wsl-ubuntu-2004   # Only LTS on WSL, so these are every 2 years
    curl.exe -L -o debian-latest.appx https://aka.ms/wsl-debian-gnulinux
    curl.exe -L -o fedora.appx https://github.com/WhitewaterFoundry/WSLFedoraRemix/releases/
Alternatively, go to the App Store and install from there:
    Ubuntu 18.04 LTS  - -  https://www.microsoft.com/store/apps/9N9TNGVNDL3Q
    Ubuntu Latest   - - -  https://www.microsoft.com/store/apps/9NBLGGH4MSV6   # Will get the latest LTS available
    OpenSUSE Leap 42  - -  https://www.microsoft.com/store/apps/9njvjts82tjx
    SUSE Linux Enterprise Server 15   https://www.microsoft.com/store/apps/9pmw35d7fnlx
    Kali Linux  - - - - -  https://www.microsoft.com/store/apps/9PKR34TNCV07
    Debian GNU/Linux  - -  https://www.microsoft.com/store/apps/9MSVKQC78PK6
    Fedora Remix for WSL   https://www.microsoft.com/store/apps/9n6gdm4k2hnc
    Pengwin   - - - - - -  https://www.microsoft.com/store/apps/9NV1GV1PXZ6P
    Pengwin Enterprise  -  https://www.microsoft.com/store/apps/9N8LP0X93VCP
    Alpine WSL  - - - - -  https://www.microsoft.com/store/apps/9p804crf0395


*** 3. Install the distro: In PowerShell, go to directory with the downloaded distro .appx file.
    Add-AppxPackage .\app_name.appx
Once your distro is installed please refer to the Initialization Steps page to initialize your new distro.
i.e. Start the distro: Open the Distro link on the Start menu, wait ~1 min for first-time initialization.
Setup a user account on first-time use.
Checking the root folder of Ubuntu 18.04, after update/upgrade it is now 878 MB
C:\Users\Boss\AppData\Local\Packages\CanonicalGroupLimited.Ubuntu18.04onWindows_79rhkp1fndgsc
The rootfs is under here where you see familiar Linux folder structure
C:\Users\Boss\AppData\Local\Packages\CanonicalGroupLimited.Ubuntu18.04onWindows_79rhkp1fndgsc\LocalState\rootfs   # For Ubuntu
C:\Users\Boss\AppData\Local\Packages\46932SUSE.openSUSELeap42.2_022rs5jcyhyac/LocalState/rootfs                   # For OpenSUSE
$ cat /etc/os-release ; uname -msr
NAME="Ubuntu", VERSION="18.04.4 LTS (Bionic Beaver)", VERSION_ID="18.04", UBUNTU_CODENAME=bionic, Linux 4.4.0-18362-Microsoft x86_64


*** 4. Updating WSL
WSL uses a standard Ubuntu installation, upgrading your packages should look very familiar:
    sudo apt-get update
    sudo apt-get upgrade
    sudo apt-get dist-upgrade


*** 5. Upgrade Ubuntu to newest distro (19.04 as of 2020-03)
Ubuntu for Windows only ever ships as the LTS version (last was 18.04 as of 2020-03).
But it is full Linux and so you can update it just like any full Linux.
https://www.how2shout.com/how-to/how-to-upgrade-ubuntu-18-04-to-19-10-on-windows-10-linux-subsystem.html

However, a few changes must be made before the upgrade.
First, we must override LTS as by default it does not allow upgrading an LTS releases to non-LTS.
Thus, we need to change this default rule to a standard one. For that type:
    cat /etc/os-release   # => 18.04 LTS
    sudo nano /etc/update-manager/release-upgrades
Change 'Prompt=lts' to 'Prompt=normal' then Ctrl-X then Y to confirm

Secondly, WSL (1 or 2) does not support LXD container technology yet and it will throw errors so must be removed.
https://askubuntu.com/questions/1119301/your-system-is-unable-to-reach-the-snap-store
LXD is a lightweight system container manager for many Linux distributions.

    sudo /etc/init.d/lxd stop
    sudo dpkg --force depends -P lxd
    sudo dpkg --force depends -P lxd-client
    sudo rm -rf /var/lib/lxd

We can now perform the upgrade 18.04 LTS => 19.04

    # sudo -S apt-mark hold procps strace sudo
    sudo -S env RELEASE_UPGRADER_NO_SCREEN=1 do-release-upgrade
    sudo apt update ; sudo apt upgrade

Finish with sudo apt upgrade

    9 installed packages are no longer supported by Canonical. You can still get support from the community.
    3 packages are going to be removed. 101 new packages are going to be installed. 422 packages are going to be upgraded.
    You have to download a total of 147 M. This download will take about 2 minutes with your connection.
    Installing the upgrade can take several hours. Once the download has finished, the process cannot be canceled.

    | Your system is unable to reach the snap store, please make sure you're connected to the Internet and update any firewall or proxy
    settings as needed so that you can reach the snap store.

    You can manually check for connectivity by running "snap info lxd"

    Aborting will cause the upgrade to fail and will require it to be re-attempted once snapd is functional on the system.

    Skipping will let the package upgrade continue but the LXD commands will not be functional until the LXD snap is installed.
    Skipping is allowed only when LXD is not activated on the system.

    Unable to reach the snap store

    sudo do-release-upgrade -c   # Check if a release upgrade is available.
Checking for a new Ubuntu release ... New release '19.10' available. ... Run 'do-release-upgrade' to upgrade to it.
    sudo do-release-upgrade      # Perform the upgrade to 19.10
Sometimes, on Windows WSL, the upgrade cannot update the latest repo available for the upgraded system. Upon checking, the
system will say there is no upgrade then manually add Ubuntu 19.10 "Eoan Ermine" official repo to our existing "Disco Dingo".
If so, sudo nano /etc/apt/sources.list and add the following line anywhere:
    deb http://archive.ubuntu.com/ubuntu/ eoan main
    
The wsl (wsl.exe) command is usable from DOS or PowerShell for all interaction with installed distros.
    wsl                 # type on its own to instantly launches the default shell. (use -d <distro name> to enter a different distro).
    wsl -e <commands>   # --exec execute commands without entering linux shell
    wsl -- <commands>   # run commands without entering linux shell

    wsl --export <Distro> <Filename>   # Exports to a tar file. Filename can be "-" for standard output.
    wsl --import <Distro> <Dest> <Filename>   # Imports the tar file as a new distribution. Filename can be "-" for standard input.
    wsl --list, -l [Options]         # List distros on this system, shows default distro with an asterisk in WSL 2.
            --all, -a                # Includes distros currently being installed or uninstalled.
            --running, -r            # List only distros that are currently running.
            --quiet, -q              # Only list the distro names, useful for scripting with grep etc. (new option in WSL 2)
            --verbose, -v            # Detailed distro info including state, shows default with an asterisk (new option in WSL 2)
    wsl --setdefault, -s <Distro>    # Sets a distro as the default.
    wsl --set-default-version <1|2>  # Set the default version that distros will install as (either verision 1 or 2). (new option in WSL 2)
    wsl --set-version <Distro> <Ver> # Convert a distro to use WSL 2 or WSL 1 architecture. e.g. wsl --set-version Ubuntu 1 (new option in WSL 2)
    wsl --terminate, -t <Distro>     # Terminates the distro. (new option in WSL 2)
    wsl --unregister <Distro>        # Unregisters the distro.
    wsl --upgrade <Distro>           # Upgrades the distribution to the WslFs file system format.
    wsl --help                       # Display usage information.
    wsl --shutdown                   # Shutdown running distros and WSL 2 lightweight utility VM (Build 18917+)
        Old WSL 1 Shutdown method was:   Get-Service LxssManager | Restart-Service
The VM that powers WSL 2 distros is only started when you need it and shut down when you don't but some cases require shutdown manually and shut down the WSL 2 VM.

Setup docker: sudo apt install docker.io
    # Need to get 63.8 MB of archives. After this operation, 319 MB of additional disk space will be used.

Setup nods.js : https://docs.microsoft.com/en-us/windows/nodejs/Install-on-wsl2

Troubleshooting & FAQ Notes:
https://docs.microsoft.com/en-us/windows/wsl/faq
https://docs.microsoft.com/en-us/windows/wsl/install-win10
https://docs.microsoft.com/en-us/windows/wsl/troubleshooting
https://www.saggiehaim.net/powershell/install-powershell-7-on-wsl-and-ubuntu/

*** What we cannot do in WSL:
Officially, no graphics interface supported so far. This means also that graphics applications cannot be executed.
To do so you have to install an unsupported X11 server such as VcXsrv or Xming.
The standard GUIs of the classic Ubuntu Linux-based are not supported for this reason of course.
The kernel of Linux is NOT part of WSL 1, but *is* part of WSL 2.

*** What I can do:
Use the command line and the basic Bash shell. It is possible to write and execute scripts.
As of Windows 10 v1803 background tasks are supported.
Develop applications (compile or cross-compile and execute them) but with no graphics so far.
Use "apt-get" to install/remove new/old packets.

*** Additional information:
other distributions are officially supported (like, for example, Debian and Kali)
these applications are free, downloadable from the Windows Store and here you can find the instruction to install and use it.
In this other question of the blog, some suggestion on how to use a GUI for WSL (unofficial, third party)

*** WSL 2
Dramatic file system performance increases, and full system call compatibility, meaning you can run more Linux apps in WSL 2 such as Docker.
WSL 2 uses an entirely new architecture that uses a real Linux kernel.
Initial builds of WSL 2 will be available through the Windows insider program by the end of June 2019.

New WSL Commands

We have also made it possible for Windows apps to access the Linux root file system (like File Explorer! Try running: explorer.exe . in the home directory of your Linux distro and see what happens) which will make this transition significantly easier.

https://github.com/microsoft/wsl/issues
https://docs.microsoft.com/en-us/windows/wsl/wsl2-ux-changes
https://devblogs.microsoft.com/commandline/announcing-wsl-2/
https://devblogs.microsoft.com/commandline/shipping-a-linux-kernel-with-windows/
https://www.omgubuntu.co.uk/2019/05/ubuntu-support-windows-subsystem-linux-2

https://askubuntu.com/questions/993225/whats-the-easiest-way-to-run-gui-apps-on-windows-subsystem-for-linux-as-of-2018
https://www.reddit.com/r/bashonubuntuonwindows/comments/fmul2g/wsl2_on_win10_with_gui_apps_under_ubuntu/
https://www.hanselman.com/blog/CoolWSLWindowsSubsystemForLinuxTipsAndTricksYouOrIDidntKnowWerePossible.aspx

*** PowerShell on WSL
You can share a Windows $Profile into WSL 2 from Windows PowerShell 5.1 without any code change:
  cp /mnt/c/Users/<<USER>>/Documents/WindowsPowerShell/Microsoft.PowerShell_profile.ps1 $Profile -Force

'@
echo $out | more
}

function Help-Timers {
    $out = @'

Techniques for timing scripts and operations (suppress output, show only the timing)
(Measure-Command { **command1; command2; etc** }).TotalSeconds

Pipe the commands to Out-Default to also show the output
(Measure-Command { ls | Out-Default }).TotalSeconds

$StartMs = (Get-Date).MilliSeconds
0..1000 | ForEach-Object { $i++ }
$EndMs = (Get-Date).MilliSeconds
Write-Host "This Script took $($EndMs - $StartMs) milliseconds to run"

https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
Measure-Command2 / Time function:   https://gist.github.com/jpoehls/2206444

Count nuber of files with a filter and time the execution
[Decimal]$Timing = ( Measure-Command {
    $Result = Get-ChildItem $env:windir\System32 -Filter *.dll -Recurse -EA SilentlyContinue
} ).TotalMilliseconds
$Timing =[math]::round($Timing)
Write-Host "TotalMilliseconds" $Timing "`nFile count:" $Result.count

$Timing = (Measure-Command {
    $Result = Get-ChildItem $env:windir\System32 -Include *.dll -Recurse -EA SilentlyContinue
} ).TotalMilliseconds
$Timing =[math]::round($Timing)
Write-Host "TotalMilliseconds" $Timing "`nFile count:" $Result.count

$StartTime = Get-Date
Start-Sleep -Seconds 5
$RunTime = New-TimeSpan -Start $StartTime -End (Get-Date)
"Execution time was {0} hours, {1} minutes, {2} seconds and {3} milliseconds." -f $RunTime.Hours,  $RunTime.Minutes,  $RunTime.Seconds,  $RunTime.Milliseconds

Using the .Net StopWatch class:
$sw = [Diagnostics.Stopwatch]::StartNew()
ls
$sw.Stop()
$sw.Elapsed.TotalSeconds

Demonstrating speed benefit of the ForEach -Parallel option
$Collection = 1..10
(Measure-Command { $Collection | ForEach-Object { Write-Host $_ }}).TotalMilliseconds                     #     4.8004 milliseconds
(Measure-Command { $Collection | ForEach-Object { Sleep 1; Write-Host $_ }}).TotalMilliseconds            # 10008.0906 milliseconds
# PowerShell 6+ only:
(Measure-Command { $Collection | ForEach-Object -Parallel { Write-Host $_ }}).TotalMilliseconds           #    40.8558 milliseconds
(Measure-Command { $Collection | ForEach-Object -Parallel { Sleep 1; Write-Host $_ }}).TotalMilliseconds  #  2095.7144 milliseconds

$Collection 1..254
(Measure-Command { $Collection | ForEach-Object { Get-WmiObject Win32_PingStatus -Filter "Address='192.168.1.$_' and Timeout=200 and ResolveAddressNames='true' and StatusCode=0" | select ProtocolAddress* } }).TotalMilliseconds
(Measure-Command { $Collection | ForEach-Object -Parallel { Get-WmiObject Win32_PingStatus -Filter "Address='192.168.1.$_' and Timeout=200 and ResolveAddressNames='true' and StatusCode=0" | select ProtocolAddress* } }).TotalMilliseconds

'@
    echo $out | more
}

function Measure-Command2 ([ScriptBlock]$Expression, [int]$Samples = 1, [Switch]$Silent, [Switch]$Long) {
<#
.SYNOPSIS
  Runs the given script block and returns the execution duration.
  Discovered on StackOverflow. http://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
  
.EXAMPLE
  Measure-Command2 { ping -n 1 google.com }
#>
  $timings = @()
  do {
    $sw = New-Object Diagnostics.Stopwatch
    if ($Silent) {
      $sw.Start()
      $null = & $Expression
      $sw.Stop()
      Write-Host "." -NoNewLine
    }
    else {
      $sw.Start()
      & $Expression
      $sw.Stop()
    }
    $timings += $sw.Elapsed
    
    $Samples--
  }
  while ($Samples -gt 0)
  
  Write-Host
  
  $stats = $timings | Measure-Object -Average -Minimum -Maximum -Property Ticks
  
  # Print the full timespan if the $Long switch was given.
  if ($Long) {  
    Write-Host "Avg: $((New-Object System.TimeSpan $stats.Average).ToString())"
    Write-Host "Min: $((New-Object System.TimeSpan $stats.Minimum).ToString())"
    Write-Host "Max: $((New-Object System.TimeSpan $stats.Maximum).ToString())"
  }
  else {
    # Otherwise just print the milliseconds which is easier to read.
    Write-Host "Avg: $((New-Object System.TimeSpan $stats.Average).TotalMilliseconds)ms"
    Write-Host "Min: $((New-Object System.TimeSpan $stats.Minimum).TotalMilliseconds)ms"
    Write-Host "Max: $((New-Object System.TimeSpan $stats.Maximum).TotalMilliseconds)ms"
  }
}

Set-Alias time Measure-Command2

function Help-PowershellSetup {
    $out = @'

choco upgrade -y PowerShell-Core   # Chocolatey for PowerShell-Core (i.e. latest pwsh 7+)
winget install powershell          # winget is built in on Windows 10+
wget https://aka.ms/install-powershell.sh; sudo bash install-powershell.sh; rm install-powershell.sh   # Linux

See here for detailed switches and usage:
https://www.thomasmaurer.ch/2019/07/how-to-install-and-update-powershell-7/

PowerShell for every system
https://github.com/proxb/PowerShell

'@
    echo $out | more
}

function Help-Timeline {
    $out = @'

Win-Tab to show Timeline

To enable Timeline in Windows 10, do the following.
    Open Settings.
    Go to Privacy - Activity History.
    Enable "Filter activities for your Microsoft Account".
    Enable the option Collect Activities. 

Timeline introduces a new way to resume past activities that you started on this PC, other Windows PCs, and iOS/Android devices. Timeline enhances Task View, allowing you to switch between currently running apps and past activities.

The default view of Timeline shows snapshots of the most relevant activities from earlier in the day or a specific past date. A new annotated scrollbar makes it easy to get back to past activities.

There's also a way to see all the activities that happened in a single day. You need to click the See all link next to the date header. Your activities will be organized into groups by hour to help you find tasks you know you worked on that morning, or whenever.

Click on the See only top activities link next to the day's header to restore the default view of Timeline.

If you can't find the activity you're looking for in the default view, search for it. There is a search box in the upper-right corner of Timeline if you can't easily locate the task you wish to restore.

Windows 10's Telemetry and Data Collection services are often being criticized by many users for collecting private or sensitive data. From their point of view, Microsoft collects too much data, especially if you are running one of the Insider Preview builds. Also, Microsoft is not transparent about what data exactly they collect, how they use it currently and what they will use it for in the future. So, this new feature might be welcomed by those who find no use for it. Probably they will be happy to disable the extra data collection option.

If you care about your privacy, you might be interested in visiting a web-based app, Microsoft Privacy Dashboard, allows the user to manage many aspects of your privacy in the new operating system. Microsoft Privacy Dashboard extends the privacy options of the built-in Settings app. While a lot of privacy options can be changed directly in Settings, they are arranged on several pages, which most users find to be inconvenient and confusing. See the following article:
https://www.howtogeek.com/fyi/windows-10-sends-your-activity-history-to-microsoft-even-if-you-tell-it-not-to/

'@
    echo $out | more
}

function Help-ToolkitQuick {
    $out = @'

### go / cc / cd   # Jump to a pre-defined jump locations
cd sys32 , cd etc , cd startup , cd temp , cd tempa , cd regstartup

These will jump you to the defined location. The actual function is 'go', but 'cc' and 'cd' are aliased to this (cd is just an alias for Set-Location in PowerShell, so can be redefined, though needs -Options AllScope to override).
Note that this can be used to jump to a registry location 'regstartup' (as the registry is a PSDrive type like "FileSystem").
Type 'go' on it's own to see the defined jump locations.
Note: By default, cd will operate normally, so if you type 'cd temp' and a subfolder caled 'temp' exists at the current location, cd will change into that directory and ignore the jump location functionality.

### def    # Show detailed information on any command type (Alias, Cmdlet, Function, ExternalScript, Application)

'@
echo $out | more
}

function Help-BeginSystemConfig { "iex ((New-Object System.Net.WebClient).DownloadString('https://bit.ly/2R7znLX'))`n# Can also run with 'iex `$(help-setup)' or 'Install-Toolkit'" }
Set-Alias Help-ToolkitSetup Help-BeginSystemConfig -Description "Show the command to start or update Custom-Tools configuration from the Github repo."
Set-Alias Help-ToolkitInstall Help-BeginSystemConfig -Description "Show the command to start or update Custom-Tools configuration from the Github repo."
Set-Alias Help-ToolkitUpdate Help-BeginSystemConfig -Description "Show the command to start or update Custom-Tools configuration from the Github repo."
Set-Alias Help-SetupToolkit Help-BeginSystemConfig -Description "Show the command to start or update Custom-Tools configuration from the Github repo."
Set-Alias Help-InstallToolkit Help-BeginSystemConfig -Description "Show the command to start or update Custom-Tools configuration from the Github repo."
Set-Alias Help-UpdateToolkit Help-BeginSystemConfig -Description "Show the command to start or update Custom-Tools configuration from the Github repo."

function Install-Toolkit {
    # Check if this function has been run dot sourced, by checking the value of $MyInvocation.InvocationName, if '.' then it was dotsourced, if 'cd' then not dotsourced
    # https://social.technet.microsoft.com/Forums/sqlserver/en-US/8e4d9f20-8479-40c1-b09f-982ab485e56e/how-to-find-out-if-a-script-is-ran-or-dotsourced?forum=winserverpowershell
    if ( $MyInvocation.InvocationName -eq 'Update-Toolkit') { "`nWarning: Command cannot run without being dotsourced! Please rerun as:`n`n   . Update-Toolkit`n" }
    else {
        echo "`niex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/roysubs/Custom-Tools/main/BeginSystemConfig.ps1'))`n"
        $answer = read-host "Would you like to reconfigure Toolkit using latest version from internet (y/n)? "
        if ($answer -eq 'y' -or $intput -eq '') {
            . iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/roysubs/Custom-Tools/main/BeginSystemConfig.ps1'))
        }
    }
}
Set-Alias Setup-Toolkit Install-Toolkit -Description "Start Custom-Tools configuration from the Github repo (or update an existing installation)."
Set-Alias Update-Toolkit Install-Toolkit -Description "Start Custom-Tools configuration from the Github repo (or update an existing installation)."

function git-push {
    if (!(Test-Path ".git")) { 
        "Crrent folder is not a git repository (no .git folder is present)" 
    }
    else { 
        "`nWill run the following if choose to continue:`n`n=>  git status  =>  git add .  [add all files]  =>  git status`n=>  git commit -m `"Update`"  =>  [pause to check]  =>  git status  =>  git push -u origin main`n`n"
        pause
        git status
        git add .
        "git status after 'git add .'"
        git status
        git commit -m "Update"
        pause
        "Note, the Github PAT (Personal Access Token) can be used in place of a password."
        git push -u origin main
    }
}

function git-push-master {
    if (!(Test-Path ".git")) { 
        "Crrent folder is not a git repository (no .git folder is present)" 
    }
    else { 
        "`nWill run the following if choose to continue:`n`n=>  git status  =>  git add .  [add all files]  =>  git status  [pause to check]`n=>  git commit -m `"Update`"    =>  git status    =>  git push -u origin master`n`n"
        pause
        git status
        git add .
        "git status after 'git add .'"
        git status
        git commit -m "Update"
        pause
        git push -u origin master
    }
}

#function Get-Subnet {
#    ####################
#    #
#    # Find all active IPs on local subnet.
#    # Discover hostnames for each.
#    # Collect shares from each and save to an output file for other scripts to use.
#    #
#    ####################
#
#    $start_time = Get-Date   # Put following lines at end of script to time it
#    # $hr = (Get-Date).Subtract($start_time).Hours ; $min = (Get-Date).Subtract($start_time).Minutes ; $sec = (Get-Date).Subtract($start_time).Seconds
#    # if ($hr -ne 0) { $times += "$hr hr " } ; if ($min -ne 0) { $times += "$min min " } ; $times += "$sec sec"
#    # "Script took $times to complete"   # $((Get-Date).Subtract($start_time).TotalSeconds)
#
#    # Use HomeFix due to use of network shares (like ING etc) to always use C:\ drive
#    # then get the name (Leaf) from $HOME in case of username aliases different from the Home folder name
#    $HomeFix = $HOME
#    $HomeLeaf = split-path $HOME -leaf   
#    if ($HomeFix -like "\\*") { $HomeFix = "C:\Users\$(Split-Path $HOME -Leaf)" }
#    $UserScriptsPath = "$HomeFix\Documents\WindowsPowerShell\Scripts"
#    if (!(Test-Path $UserScriptsPath)) { md $UserScriptsPath -Force -EA silent | Out-Null }
#    # Get running script name and build the external filenames
#    # https://stackoverflow.com/questions/817198/how-can-i-get-the-current-powershell-executing-file
#    # $ScriptFile = gci $MyInvocation.MyCommand.Path
#    # $ScriptFull = $ScriptFile.FullName
#    # $ScriptPath = Split-Path $ScriptFull -Parent
#    # $ScriptName = Split-Path $ScriptFull -Leaf
#    # $ScriptExt  = $ScriptFile.Extension
#    # $ScriptNoExt = ($ScriptName -Split $ScriptExt)[0]       # Always use -Split for a string; .split('xyz') will split on that array of characters
#    # $ScriptPWD = Join-Path $ScriptPath "$ScriptNoExt.pwd"   # Known user/pass details, store in format: Hostname:::Username:::Password, one per line
#    # $OutIPs = Join-Path $ScriptPath "$ScriptNoExt.ip"          # Output: store the discovered IPs
#    # $OutShares = Join-Path $ScriptPath "$ScriptNoExt.shares"   # Output: store found shares to use for other scripts
#    $OutIPs = "$env:TEMP\Get-Subnet.ip"
#    $OutShares = "$env:TEMP\Get-Subnet.shares"
#
#    # Does Ping-IPRange.ps1 exist, if not, get it from PS Gallery
#    # Also, add the Users Scripts folder to the PATH if not already on the User PATH statement
#    if (!(Test-Path "$UserScriptsPath\Ping-IPRange.ps1")) { 
#
#        $url = 'https://gallery.technet.microsoft.com/scriptcenter/Fast-asynchronous-ping-IP-d0a5cf0e/file/124575/1/Ping-IPrange.ps1'
#        $FileName = ($url -split "/")[-1]   # Could also use:  $url -split "/" | select -last 1   # 'hi there, how are you' -split '\s+' | select -last 1
#        $OutPath = Join-Path $UserScriptsPath $FileName
#        Write-Host "Downloading  $FileName to $OutPath ..."
#        try { (New-Object System.Net.WebClient).DownloadString($url) | Out-File $OutPath }
#        catch { "Could not download $FileName ..." }
#
#        $RegistryUserPath = "HKCU:\Environment"
#        $PathOld = (Get-ItemProperty -Path $RegistryUserPath -Name PATH).Path
#        $PathArray = $PathOld -Split ";" -replace "\\+$", ""
#        # Add to Path if required
#        $FoundPath = 0
#        foreach ($Path in $PathArray) { 
#            if ($Path -contains $UserScriptsPath ) { $FoundPath = 1 } 
#        }
#        if ($FoundPath -eq 0) {
#            $PathNew = $PathOld + ";" + $UserScriptsPath
#            Set-ItemProperty -Path $RegistryUserPath -Name PATH -Value $PathNew
#            (Get-ItemProperty -Path $RegistryUserPath -Name PATH).Path
#        }
#    }
#
#    Function Test-CommandExists($command) {
#        $oldPreference = $ErrorActionPreference
#        $ErrorActionPreference = 'Stop'
#        try { if (Get-Command $command) { return $true } }       # Note that return is not required here, but handy to clarify that the function will return this
#        catch { Write-Host "$command does not exist" ; return $false }
#        finally { $ErrorActionPreference = $oldPreference }    
#    }
#
#    # Test if the "Ping-IPRange" Function is available. If not, then run the script to load the function
#    if (!(Test-Path "$UserScriptsPath\Ping-IPRange.ps1")) { "Ping-IPRange.ps1 script is not available, download must have failed." ; exit }
#    . $UserScriptsPath\Ping-IPRange.ps1
#    if (!(Test-CommandExists Ping-IPRange)) { "Ping-IPRange function did not load from Ping-IPRange.ps1 script." ; exit }
#
#    # To establish the range to check, get the Enabled IPV4 addresses
#    # From these, we should pick the one starting "192.168" (an assumption, but works for most home networks, 10.0.0. is sometimes used also)
#    $ipenabledarray = (Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | ? { $_.IPEnabled -eq $true } | select -ExpandProperty IPAddress)
#    # Looking for 192.168.0.x or 192.168.1.x
#    $iphere = $ipenabledarray | ? { $_ -match "192\.168\.[0-1]\." }      # -match "10\.0" could also work for home networks
#    $subnet = ($iphere -split '\.')[0,1,2] -join "."
#    $ipstart = "$subnet.1"
#    $ipend= "$subnet.255"
#    # "$ipenabledarray`n$iphere`n$subnet`n$ipstart`n$ipend`n"
#
#    # $subnet = ((Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | ? { $_.IPEnabled -eq $true } | select -ExpandProperty IPAddress)[0] -split '\.')[0,1,2] -join "."
#    ""
#    $shares = @()
#    $localdrives = @()
#
#    # Now can run the script to collect live hosts
#    # Method 1: Use a range,   Ping-IPRange 192.168.1.1 192.168.1.255
#    # Method 2: If a previous run happened, use that info to just ping between that range.
#
#    # Check for file with existing IPs and only scan that range on this run
#    $iparray = @()
#    if (Test-Path $OutIPs) { 
#        $ipstart = Get-Content $OutIPs -First 1
#        $ipend = Get-Content $OutIPs -Last 1    # can also use say: last 3 lines with [-1 .. -3], but more for characters than lines
#        if ($null -eq $ipstart -or $null -eq $ipend) { "Problem with IP file format: $OutIPs" ; break }
#    }
#
#    ####################
#    #
#    # Viewing and resolving shares on remote systems
#    #
#    ####################
#    # Using CIM throws WinRM errors. Could use if WinRM is configured on all clients, but for now ignore this and use net view as it is always available
#    # $cim = New-CimSession -ComputerName $hostname
#    # $sharescim = Get-SmbShare -CimSession $cim
#
#    # net view has quirks and issues, but can be made to work quite well.
#    # Firstly, found that 'net view' on its own, on Windows 10, almost always results in an error:
#    #    "System error 6118 has occurred. The list of servers for this workgroup is not currently available."
#    # This results *even* when I enable all of the points in the below on *all* hosts, so this was a dead-end to some extent
#    # e.g. "Windows Feature > SMB 1.0/CIFS File Sharing Support, make sure this is enabled and reboot." on all hosts does not help.
#    # https://superuser.com/questions/1347496/windows-10-cmd-net-view-all-returns-error-6118-list-of-servers-not-available
#    # https://social.technet.microsoft.com/Forums/windows/en-US/b8a92fb2-3ed5-4785-bd8f-990a1c54c1c2/net-view-command-not-working-after-upgrading-our-workstations-to-windows-10-1703-creators-edition
#    # https://docs.microsoft.com/en-au/windows-server/storage/file-server/troubleshoot/smbv1-not-installed-by-default-in-windows
#    # https://systemhelpsolution.blogspot.com/2011/01/system-error-6118-has-occurred-list-of.html
#    # https://social.technet.microsoft.com/Forums/en-US/32ea7c07-ac85-43a7-a334-88f8970b2e19/network-problem-net-view-fails-error-6118-net-view-server-successfully-lists-shares?forum=w7itpronetworking
#    # https://www.tenforums.com/network-sharing/136473-error-6118-net-view-problem-solved-last.html?__cf_chl_jschl_tk__=7a62541f1ad661b0c91c373a2b7f584d543b4257-1601906190-0-Ace9mAbL36tAFIdsK6ceMnQJwVHlLUicHn35-8ghYQBl0bNCHnTmm4qYr-05qAZ5GRyLgzgN69lDTlCuH-yc_HO3EbTmbYP5R7tTeJi_CLJBonFTnPkNPYAhqmVH3sZyn_jyPFkgpQiImC9GE45MXo9n5ovIq0kly74KcmblN4jfBSHSCKY2ON4Uljgpc5APRFPf022C73yVHBpGz2wk6G0_PCFg7D5pOyL7nBT6GQhzv5YVhD4vRwlGTy2aW-ZNplVl886u1mI8IDXPnl0fxyfQrwSt8KqlmI3n3ThLeX3YCiU5fVkGZBAEnVom5YejzXO6inaMCLHCtndURfD3ITIGYglaZQoFOVkO52tRtItyS9HYXWGxDGV6Nve8uAF9IqskANVPkwTh_DnlbTDj7uin9fD18x8ZDlWcRxjVCidB
#
#    # However, as is often the case, the solution is quite simple when using the right tools:
#    # cmdkey /list   # or, go to Credential Manager for the same information with:   control keymgr.dll
#    # Net View will fail for an unknown host:  # net view <hostname>   =>   "System error 5 has occurred. Access is denied."
#    # Add the host credentials with:           # cmdkey /add:<hostname> /user:<username> /pass:<password>
#    # After this, can now use net view         # net view <hostname>   =>   Shows list of all shares
#    # Remove existing credentials with:        # cmdkey /delete:<hostname>
#    # Note that after removing the credentials, net view will continue to work until the next reboot.
#
#    # Check for file with existing Hostname#Username#Password (using ":::" as delimiter, which assumes that 3x ":" in a row are not in password)
#    $hosts = @()
#    if (Test-Path $ScriptPWD) { 
#        foreach ($i in gc $ScriptPWD) { $hosts += $i }
#        # Assume all credentials in the .pwd file are correct, so remove and then reload credentials for each host
#        foreach ($h in $hosts) {
#            $hostname = ($h -split ':::')[0]
#            $username = ($h -split ':::')[1]
#            $password = ($h -split ':::')[2]
#            $cmdkeydel = cmdkey /delete:$hostname 2>$null
#            $cmdkeyadd = cmdkey /add:$hostname /user:$username /pass:$password 2>$null
#            # net use \\$hostname\ipc$ /user:$username $password
#        }
#    }
#
#    # Notes on optional features ... may need to install SMB 1.0 for some of this:
#    # Get-WindowsOptionalFeature -Online -FeatureName *smb*
#
#    $shares = @()
#    # ping each ip in the range and use that to resolve all 
#    foreach ($i in $(ping-iprange $ipstart $ipend | select ipaddress)) {
#
#        try { $ip = $i.IPAddress } catch { $ip = "NA" }
#        if ($ip -ne "NA") { $iparray += $ip.IPAddressToString }
#        try { $hostname = [System.Net.Dns]::GetHostByAddress($i.ipaddress).Hostname } catch { $hostname = "NA" }
#        # The above method attaches ".Home" to 192.168.1.1 and to the local system, so remove that
#        $hostname = ($hostname -split "\.Home")[0]
#
#        # Collect local drives for the local computer
#        if ($i.IPAddress -eq $iphere) {
#            foreach ($drive in $(Get-WmiObject -query "SELECT * from win32_logicaldisk where DriveType = '3'" | select DeviceID)) {
#                if ($drive.DeviceID -ne "C:") {
#                    $localdrives += $drive.DeviceID
#                }
#            }
#        }
#    }
##
##        $netview = "No shares found"
##        if ($hostname -ne "NA") {      # -and $i.IPAddress -ne $iphere
##            $netview = @()   # initialise to ensure 
##            # I asked the following on external (non-PowerShell) outputs: https://stackoverflow.com/questions/64196373/suppress-and-handle-stderr-error-output-in-powershell-script?noredirect=1#comment113520006_64196373
##            # As an unfortunate side effect, in versions up to v7.0, a 2> redirection can also throw a script-terminating error if $ErrorActionPreference = 'Stop' happens to be in effect, if at least one stderr line is emitted.
##            $netview = net view \\$hostname /all 2>$null | select -Skip 7 | ? {$_ -match 'disk*'} | % { $_ -match '^(.+?)\s+Disk*' | Out-Null ; $matches[1] }
##            if ([string]::IsNullOrWhiteSpace($netview)) { $netview = "No shares found" }
##            $cmdkey = cmdkey /add:$hostname /User:Boss /Pass:boss 2>$null
##            # if ($null -eq $netview)   # https://thinkpowershell.com/test-powershell-variable-for-null-empty-string-and-white-space/
##        }
##        # https://www.itprotoday.com/powershell/view-all-shares-remote-machine-powershell
##        #   | % ForEach-Object { $_.Matches.Value }
##        #   | Select-String $reg -AllMatches | % ForEach-Object { $_.Matches.Value }
##
##        # Ok, big breakthrough on this, don't know which step cracked it but do know that the security login was key to the shares appearing.
##        # Note that the shares themselves are completely non-password protected, so anyone should be able to access the shares, but that was not in fact happening.
##        # HP1 and HP3 shares were completely invisible to me until I physically established the connection via Network > HP1 > and entering credentials to view the shares.
##        # As soon as I did this, all shares are now visible on each of those hosts.
##        # Presumably, the following can save credeintials (using IPC$ just to establish connection, and the /savecred switch)
##        # net use \\server\IPC$ <password> /savecred /user:Boss /persistent:yes
##        # net use * \\server\share /user:domain\user password /persistent:yes
##        # cmdkey /add:\\server /user:username /pass:password
##        # https://social.technet.microsoft.com/Forums/scriptcenter/en-US/8fc5d667-c1f9-4670-a6e1-3ba2c5e8603d/net-use-savcred-user-conflicting-switches
##
##        # Correct Answer:
##        # /savecred option is a dead end if you want to use with a AD connection and specify the user.
##        # Instead, add something to credential manager and then windows will auto-magicly use it to logon
##        # to a persistent share on next logon.
##        #   cmdkey /add:%Server% /User:%domain%\%user% /Pass:%password%
##        # Then, mount a share or try to access the server (net view for example). The saved credentials will be used:
##        #   net use P: \\%Server%\share /user:%domain%\%user% /persistent:yes
##        # You can check stored credentials by typing
##        # After this, shares will work persistently after reboots etc.    
##        # ToDo: look for PowerShell equivalents of the above commands.
##
##        # 1. If net view does not work, open a dialogue to collect user/pass information for that server.
##        # 2. cmdkey /add:%Server% /User:%domain%\%user% /Pass:%password%
##        #    Possibility to load the credentials from a text file so that they are always available, to reuse later
##        # Note: no need for IPC$ or other workarounds like that, once the correct credentials are stored, the servers will always be accessible.
##
##        "$ip : $hostname : $netview"
##
##        if ($netview -ne "No shares found") {
##            foreach ($m in $netview) { 
##                # Using hidden ($) shares would be best, but that would require an admin password for each host (or pass-through authentication if password here is same as there)
##                # Instead of that, just look for the open shares that are available without password and crawl those (so exclude C: drive, all hidden, and non-useful entry points)
##                # or main media analysis, following is non-generic / my share conventions. Also remove all hidden ($) shares: C$, D$, "Drives$" (using $ regex end of string)"
##                # if ($m -notmatch "\$" -and $m -notmatch "Drives$" -and $m -notmatch "Drive-C" -and $m -notmatch "Downloads-C" -and $m -notmatch "Downloads-D" -and $m -notmatch "Desktop-C") {
##                if ($m -match "Drive\-" -and $m -notmatch "Drive\-C") {
##                    $shares += "\\$hostname\$m"
##                }
##            }
##        }
##    }
##
##    # https://morgantechspace.com/2015/06/powershell-find-machine-name-from-ip-address.html
##    # [System.Net.Dns]::GetHostByAddress($ipAddress).Hostname
##    # Python notes:
##    # https://www.tutorialspoint.com/python_penetration_testing/python_penetration_testing_network_scanner.htm
##
##    ""
##    "Array of drives on '$(hostname)' local computer (excluding C:):"
##    $localdrives   # no need to sort 
##    ""
##    "Array of required network shares (root drives only; exclude C: and hidden shares):"
##    $shares = $shares | sort   # sort shares before display and outputting to share file
##    $shares   # no need to sort these 
##    ""
##    $hr = (Get-Date).Subtract($start_time).Hours ; $min = (Get-Date).Subtract($start_time).Minutes ; $sec = (Get-Date).Subtract($start_time).Seconds
##    if ($hr -ne 0) { $times += "$hr hr " } ; if ($min -ne 0) { $times += "$min min " } ; $times += "$sec sec"
##    "Script took $times to complete"   # $((Get-Date).Subtract($start_time).TotalSeconds)   # "Elapsed Time: $(($_time-$startDTM).totalseconds) seconds"
##    ""
##    rm $OutIPs -Force -ErrorAction SilentlyContinue
##    foreach ($i in $iparray) { Add-Content $OutIPs $i }   # Add each discovered IP to the output file
##
##    rm $OutShares -Force -ErrorAction SilentlyContinue
##    foreach ($i in $shares) { Add-Content $OutShares $i }   # Add each discovered share to the output file
#}

########################################
#
# Sample Code (does not run): Template input from TXT and CSV. Output to CSV
# Leave uncommented for readability (put "Exit" above this section so that it never runs)
# To add: CliXml code examples.
#
#########################################

# # Get IP addresses from each line of text file
# Get-Content C:\0\ip-addresses.txt | ForEach-Object {
#     $hostname = ([System.Net.Dns]::GetHostByAddress($_)).Hostname
#     if ($? -eq $True) { $_ +": "+ $hostname >> "C:\machinenames.txt" }
#     else { $_ +": Cannot resolve hostname" >> "C:\machinenames.txt" }
# }
# 
# # Get IP addresses from each line of CSV file, and output IP and Hostname to a different CSV
# Import-Csv C:\0\ip-addresses.csv | ForEach-Object {
#     $hostname = ([System.Net.Dns]::GetHostByAddress($_.IPAddress)).Hostname
#     if ($? -eq $False) { $hostname="Cannot resolve hostname" }
#     New-Object -TypeName PSObject -Property @{
#         IPAddress = $_.IPAddress
#         HostName = $hostname
#     }
# } | Export-Csv C:\0\machinenames.csv -NoTypeInformation -Encoding UTF8





function Help-Basics {
    $out = @'

:: PowerShell General Basics ::

help about_*   # display all built-in long-form documentation
help *run*     # display all help files for commands with *run* in the name

:: Setting values vs comparison operations (the biggest 'gotcha' in all of PowerShell!).
    $x = 10       # Sets $x to be equal to 10
    if ($x = 5)   # This is wrong! This does NOT test if $x is equal to 5 (!!)
The above statement will always complete without error and return $true but instead of checking
if $x equal 5, it just sets $x to 5 internal to the 'if' test, which then always returns '$true'.
$x could have been some other value and this resets it to 5. The correct way is:
    if ($x -eq 5)
Almost never use "=" inside an if statement. The equivalence operators are:
-eq, -ne, -gt, -lt, -ge, -le  (eq=equalto, n=not, e=equal, g=greater, l=less, t=than)
-Like, -NotLike, -Match, -NotMatch, -Contains, -NotContains, -In, -NotIn, -Replace
help about_Comparison_Operators   (help about_compar*)
help about_Assignment_Operators   (help about_assign*)

:: You can express numbers with mb / gb values:
    $x = 5mb ;   $x will then reutrn '5242880'.

:: Wildcards in paths:
dir "c:\Pro*\*Pow*\Mod????\*"
help about_Wildcards for more   (help about_wild*)

:: Showing the pipeline and extract Property information from Objects
Get-ChildItem -Path c:\Windows | Where-Object {$_ -like 'Win*'} | ForEach-Object { Write-Output -InputObject "$($_.Name) : $($_.LastWriteTime)" }
ls c:\windows | ? {$_ -like 'Win*'} | % {echo "$_.Name : $_.LastWriteTime"}
# Way to remember the '?' / '%' aliases: 'Where' (?) is a question, while ForEach (%) is not a question!
help about_Pipelines    (help about_pip*)

:: Looking at the aliases in the above:
Get-Alias -Name ls,%,?,echo      # 
# Note how h and r also return here, due to '?' being a wildcard here
# For that with: Get-Alias -Name ls,%,'`?',echo
A note on aliases: People often say "never use aliases in scripts!". This is a *lot of crap*.
PowerShell has existed since 2006, and in 13+ years, the core aliases have *never* changed so
this should be revised to "Use core aliases in scripts always, but never use *custom* aliases
in scripts (unless you define the alias inside the script)". i.e. gci, ?, %, sls, foreach, select
are all great to use and can make scripts massively easier to parse over long-winded lines.
help about_aliases

:: Common Parameters
-Debug (-DB), -ErrorAction (-EA), -ErrorVariable (-EV), -InformationAction, -InformationVariable
-OutVariable (-OV), -OutBuffer (-OB), -PipelineVariable (-PV), -Verbose (-VB),
-WarningAction (-WA), -WarningVariable (-WV)
Risk mitigation parameters: -WhatIf (-WI), -Confirm (-CF)
help about_CommonParameters   (help about_common*)

:: History
h (history, list history), r (invoke-history, run history command) (see Id and Count parameters)
h | ? {$_.CommandLine -like "*Service*" }
h | fl -Property *
h | Export-Csv History.csv
#<part-of-commnand> then Tab: you must put # as first character, then part of the command in history, then Tab
F8 / Shift-F8 : back / forwards through history (you can type part of a command to start)
(DosKey) F7 : View history overlay (removed in latest Win 10 I think, though Win 7 and early Win 10 had this)
(DosKey) Alt+F7 : Clear the command history (removed in latest Win 10)
:: PSReadLine extends console functionality (standard component in Win 10 but not Win 7 / 8):
Ctrl+r   # Reverse search in history (Ctrl+s is forward search, little used) (Do not work in ISE).
For Win 7 / 8 : Install-Module PSReadLine
Get-PSReadLineKeyHandler | ? {$_.function -like '*hist*'}  # Show features.
# Get-PSReadLineOption to show values and Set-PSReadLineOptions to set them
Get-PSReadlineOption   # http://woshub.com/powershell-commands-history/
Remove-Item (Get-PSReadlineOption).HistorySavePath   # Remove all history by deleting the save file
https://github.com/PowerShell/PSReadLine
help about_History

:: Common Get-Date outputs:
"$(Get-Date -format g) Start logging"    2/5/2008 9:15 PM
"$(Get-Date -format F) Start logging"    Tuesday, February 05, 2008 9:15:13 PM
"$(Get-Date -format o) Start logging"    2008-02-05T21:15:13.0368750-05:00
"$(Get-Date -format "yyyy-MM-dd__HH-mm"  Note HH for 24-hr and hh for 12-hr time

:: Automatic Variables
$? ($LastExitCode) contains execution status of last operation ($true if success or $false if failed).
$_ ($PSItem) contains current object in the pipeline object.
$^ / $$ cntains the first / last token in the last line received by the session.
$ForEach contains enumerator of a ForEach loop.   help about_ForEach
$AllNodes / $Args / $ConsoleFileName / $Error / $Event / $EventArgs / $EventSubscriber / $ExecutionContext
$Home / $Profile / $Host / $MyInvocation / $NestedPromptLevel
$NULL / $OFS / $PID / $PSBoundParameters / $PSCmdlet / $PSCommandPath / $PSCulture / $PSDebugContext
$PsHome / $PsScriptRoot / $PsSenderInfo / PsUICulture / $PsVersionTable / $Pwd / $Sender
$ReportErrorShowStackTrace / $ShellID / $StackTrace / $This / $True / $False
/ $Input / $LastExitCode / $Matches /


'@
echo $out | more
# https://tommymaynard.com/cmdlet-and-function-alias-best-practice-2015/
}

function Help-Edit {
    $out = @'

:: Line Editing provided by ReadLine Module ::
::::::::::::::::::::::::::::::::::::::::::::::

These functions are standard on PS v5+, for previous versions of PS, you must install ReadLine:
Install-Module ReadLine
    
1. ` [Back apostrophe key] : Insert a line break to continue longer commands onto next line.
     You do not need this if you have a pipe (|) at the end of a line (assumes continuation).
2. ` Alternatively, inside a string, this is an escape character to make a literal character.

:: Edit current line
Left / Right arrow : Move left / right on the current line.
Up / Down arrow : Move back / forward in command history (current line will remain in buffer)
Page Up / Page Down : Scroll screen buffer up / down (current line will remain in buffer)
Home / End : Move to the beginning or end of the line.
Insert : Normal function (switch between insert mode and overwrite mode).
Tab / Shift+Tab Press the Tab key or press Shift+Tab to access the tab expansion function.
Esc : Clear the current line
Ctrl+C : Break out of the subprompt or terminate execution.
Ctrl+S Press Ctrl+S to pause or resume the display of output.
Alt + Mouse drag : select a block region in the PowerShell console.

:: Quick movement and delete
Ctrl+Left arrow / Ctrl+Right arrow : to move left or right one word at a time.
Ctrl+End : Delete all the characters in the line after the cursor.
Ctrl+Backspace : Delete from current cursor position left until the start of the current word.
Ctrl+Delete : Delete from current cursor position right until the end of the current word.

:: Right-click : If QuickEdit is disabled, displays an editing shortcut menu with
Mark, Copy, Paste, Select All, Scroll, and Find options.
To copy the screen buffer to the Clipboard, right-click, choose Select, and then press Enter.

:: PowerShell Console menu
Alt+Space then E : Display the editing shortcut menu (Mark, Copy, Paste, Select All, Scroll, and Find options)
Cltr+K (Mark), Ctrl+Y (Copy), Ctrl+P (Paste), Ctrl+S (Select All),
Ctrl+L (scroll through screen buffer), Ctrl+F (search for text in the screen buffer).
To copy the screen buffer to the Clipboard, use: Alt+Space then E then S, then press Alt+Space then E then Y.

> F1 Moves the cursor one character to the right on the command line. At the end of the line, inserts one character from the text of your last command.
> F2 Creates a new command line by copying your last command line up to the character you type.
> F3 Completes the command line with the content from your last command line, starting from the current cursor position to the end of the line.
> F4 Deletes characters from your current command line, starting from the current cursor position up to the character you type.
> F5 Scans backward through your command history.
> F7 Displays a pop-up window with your command history and allows you to select a command. Use the arrow keys to scroll through the list. Press Enter to select a command to run, or press the Right arrow key to place the text on the command line.
> F8 With text entered on the prompt, press F8 to scan back through command history for commands that match the text typed so far.
> F9 Runs a specific numbered command from your command history. Command numbers are listed when you press F7.

:: History:
h (history, list history), r (invoke-history, run history command)
Ctrl+r (reverse-search-command history), Ctrl-s (forward-search-command-history) - These do not work in ISE
#<part-of-commnand> then Tab: you must put # as first character, then part of the command in history, then Tab
F8 / Shift-F8 : back / forwards through history (you can type part of a command to start)
(DosKey) F7 : View history overlay (removed in latest Win 10 I think, though Win 7 and early Win 10 had this)
(DosKey) Alt+F7 : Clear the command history (removed in latest Win 10)

'@
echo $out | more
}

function Help-Truncate {
    $out = @'

:: Truncated PowerShell Output ::
Sometimes PowerShell truncates output, and it can be very confusing if you are expecting more output,
PowerShell replaces that with an ellipsis, taunting you...

Column Width: If it's just a column width problem, just pipe to Out-String with the -Width parameter.

> BEFORE (Note that 17 items truncated to 4)
Get-Module -ListAvailable | group version | sort Name -Descending

Count Name                      Group
----- ----                      -----
    4 3.0.0.0                   {Microsoft.PowerShell.Diagnostics, Microsoft.PowerShell.Host, Microsoft.PowerShell.S...
    5 2.12.0                    {Boxstarter.Bootstrapper, Boxstarter.Chocolatey, Boxstarter.Common, Boxstarter.Hyper...
   17 2.0.0.0                   {AppLocker, BitsTransfer, Hyper-V, International...}

> AFTER (note that more values are displayed but still not everything)
Get-Module -ListAvailable | group version | sort Name -Descending | Out-String -Width 160

So this doesn't fix it ...

The object might be an array/collection, and PowerShell is only showing the first few entries in that array, rather than the lot.

The fix is to change the $FormatEnumerationLimit value. Type it on its own:
$FormatEnumerationLimit
Previous versions were "3" and most recent Win 10 is "4". Set to "-1" to show all values.

$FormatEnumerationLimit=-1

> AFTER (note that more values are displayed, without using an Out-String -Width)
Get-Module -ListAvailable | group version | sort Name -Descending

Count Name                      Group
----- ----                      -----
    4 3.0.0.0                   {Microsoft.PowerShell.Diagnostics, Microsoft.PowerShell.Host, Microsoft.PowerShell.S...
    5 2.12.0                    {Boxstarter.Bootstrapper, Boxstarter.Chocolatey, Boxstarter.Common, Boxstarter.Hyper...
   17 2.0.0.0                   {AppLocker, BitsTransfer, Hyper-V, International, NetAdapter, NetLbfo, NetQos, NetSe...


Still doesn't fix the problem as now truncating by width of screen!

The fix is to use a much larger value for 'Out-String -Width xxx', but we must also use
'Format-Table -Property * -AutoSize' to minimize column width...

Get-Module -ListAvailable | group version | sort Name -Descending | Format-Table -Property * -AutoSize | Out-String -Width 4096

This looks pretty ugly on-screen so could just get the Group that we want (the one with 17) in various
ways by selecting just that item:

Get-Module -ListAvailable | group version | ? { $_.Name -eq '2.0.0.0' } | Select Group | Format-Table -Property * -AutoSize | Out-String -Width 4096

This then displays all 17 values in the Group (as long as the $FormatEnumerationLimit is set to -1)

Another approach is to just pipe it to format-table -wrap, which will show the lot + reveal long fields:
Get-Module -ListAvailable | group version | ? { $_.Name -eq '2.0.0.0' } | format-table -wrap

https://greiginsydney.com/viewing-truncated-powershell-output/

'@
echo $out | more
}

function Help-DotNetStringOperators {
    $out = @'

:: The String operators .replace, .split, .trim etc

These are not part of PowerShell and are instead part of the .NET System.String class.
For that reason, they do not have built-in documentation.

However, you can get some information are follows:

$i = 'mytemp'     # just create a dummy variable
$i | Get-Member   # This will show you all of the possible methods
$i.Split          # Picking a System.String method, we can see the OverloadDefinitions

OverloadDefinitions
-------------------
string[] Split(Params char[] separator)
string[] Split(char[] separator, int count)
string[] Split(char[] separator, System.StringSplitOptions options)
string[] Split(char[] separator, int count, System.StringSplitOptions options)
string[] Split(string[] separator, System.StringSplitOptions options)
string[] Split(string[] separator, int count, System.StringSplitOptions options)

There are still some problems with this, since we cannot see what valid options exist
for things like "Sysmte.StringSplitOptions options" but the following helps:

[Enum]::GetNames([StringSplitOptions]) 

-------------------
None
RemoveEmptyEntries

We can also google on those specific names StringSplitOptions for more info.

You can also iterate over all of the above methods with something like:
"mystring" | Get-Member | ? { $_.MemberType -eq "Method" } | % { ":: $($_).Name:" ; $i.($_.Name) }

'@
echo $out | more
}
    
function Help-Output {
    $out = @'

Firstly, the purist PowerShell position (from the author of PowerShell even!) is "Never
use Write-Host. If you do, then you are doing it wrong!". But he only cares about the 
"Holy Pipeline". For most normal people, the pipeline is not the be all and end all.
https://www.jsnover.com/blog/2013/12/07/write-host-considered-harmful/

Here is the problem:

Write-Host '------- 1' ; Get-Item / | Select-Object FullName ; Write-Host '------- 2'

If you run this, you will see that the two Write-Host's both appear on screen first
followed by the Get-Item output. That is roughly because the Write-Hosts are thrown out
as soon as they are encountered, but the pipeline enabled line is processing in a
different way.

If you want to fix this (and you don't care about Snover's purist pipeline stuff), just
put " | Out-Host" on the end of pipeline lines that are causing problems. For writing
something that will be used in production you can remove all Write-Host and Out-Host
things and make scripts fully pipeline enabled and write information to log files
instead. For now, I'll just keep using Write-Host and then " | Out-Host" a few of the
pipeline lines that are causing problems to make everything sequential. i.e.

Write-Host '------- 1'
Get-Item / | Select-Object FullName | Out-Host
Write-Host '------- 2'

'@
    echo $out | more
}

function Help-Robocopy {
    $out = @'

/XJD :: eXclude Junction points and symbolic links for Directories.
/XJF :: eXclude symbolic links for Files.
/XJ  :: eXclude Junction points and symbolic links. (normally included by default).
/SL  :: copy symbolic links versus the target.
/R:n :: number of Retries on failed copies: default 1 million.
/W:n :: Wait time between retries: default is 30 seconds.

/TIMFIX  :: FIX file TIMes on all files, even skipped files.
/FFT     :: assume FAT File Times (2-second granularity).
/DST     :: compensate for one-hour DST time differences.
/ZB      :: use restartable mode; if access denied use Backup mode.
/J       :: copy using unbuffered I/O (recommended for large files).
/COPYALL :: COPY ALL file info (equivalent to /COPY:DATSOU).
/SECFIX  :: FIX file SECurity on all files, even skipped files.
/SEC     :: copy files with SECurity (equivalent to /COPY:DATS).
/PURGE   :: delete dest files/dirs that no longer exist in source.
/MIR     :: MIRror a directory tree (equivalent to /E plus /PURGE).
/MOV     :: MOVe files (delete from source after copying).
/MOVE    :: MOVE files AND dirs (delete from source after copying).
/CREATE  :: CREATE directory tree and zero-length files only.
/MON:n   :: MONitor source; run again when more than n changes seen.
/MOT:m   :: MOnitor source; run again in m minutes Time, if changed.
/RH:hhmm-hhmm :: Run Hours - times when new copies may be started.
/PF      :: check run hours on a Per File (not per pass) basis.
/IPG:n   :: Inter-Packet Gap (ms), to free bandwidth on slow lines.
/MT[:n]  :: Do multi-threaded copies with n threads (default 8).
            n must be at least 1 and not greater than 128.
            This option is incompatible with the /IPG and /EFSRAW options.
            Redirect output using /LOG option for better performance.
/DCOPY:copyflag[s] :: what to COPY for directories (default is /DCOPY:DA).
          (copyflags : D=Data, A=Attributes, T=Timestamps).
/NODCOPY :: COPY NO directory info (by default /DCOPY:DA is done).
/NOOFFLOAD :: copy files without using the Windows Copy Offload mechanism.

/XF file [file]... :: eXclude Files matching given names/paths/wildcards.
/XD dirs [dirs]... :: eXclude Directories matching given names/paths.
/IA:[RASHCNETO] :: Include only files with any of the given Attributes set.
/XA:[RASHCNETO] :: eXclude files with any of the given Attributes set.
   /XC :: eXclude Changed files.
   /XN :: eXclude Newer files.
   /XO :: eXclude Older files.
   /XX :: eXclude eXtra files and directories.
   /XL :: eXclude Lonely files and directories.
   /IS :: Include Same files.
   /IT :: Include Tweaked files.
/MAX:n :: MAXimum file size - exclude files bigger than n bytes.
/MIN:n :: MINimum file size - exclude files smaller than n bytes.
/MAXAGE:n :: MAXimum file AGE - exclude files older than n days/date.
/MINAGE:n :: MINimum file AGE - exclude files newer than n days/date.
/MAXLAD:n :: MAXimum Last Access Date - exclude files unused since n.
/MINLAD:n :: MINimum Last Access Date - exclude files used since n.
             (If n < 1900 then n = n days, else n = YYYYMMDD date).
    /L    :: List only - don't copy, timestamp or delete any files.
    /X    :: report all eXtra files, not just those selected.
    /V    :: produce Verbose output, showing skipped files.
  /ETA    :: show Estimated Time of Arrival of copied files.

Note: Using /PURGE or /MIR on root of the volume formerly applied to System
Volume Information. This is no longer the case; robocopy will skip any files
or directories with that name in the top-level source and destination
directories of the copy session.

Robocopy notes:
# Save to Registry as default settings: robocopy /REG /R:1 /W:1                    

# In most cases, remember the following most used defaults:
robocopy <src> <dst> /r:1 /w:1 /mir /xjd /xd $recycle.bin

'@
echo $out
}



function Help-sls {
    $out = @'
    
:: sls (Select-String) vs grep vs findstr

grep: grep <pattern> files.txt   cat *.log | grep <pattern>
sls:  sls <pattern> files.txt    cat *.log | sls <pattern>
Note: the default parameter position for sls is pattern first, then file.

Note that all PowerShell Cmdlets are case-insensitive, while all Linux tools are
case-sensitive. Also note that older tools like findstr.exe are case-sensitive too.
Have to keep this in mind when using these tools.

findstr can be used in PowerShell, and works in situations where sls does not:
Get-Service | findstr "Sec"  << works fine
Get-Service | sls "Sec"      << returns nothing!
The sls returns nothing here because sls works with strings, but Get-Service here
returns a pipeline object, while findstr can only see that output as a string.

Out-String can be used to turn a piped input object (an array of service objects
in the case of 'Get-Service') into a single string, and the -Stream switch allows
each line to be output as a single string instead. sls also supports the
-CaseSensitive switch to change case sensitivity.

# Case-insensitive regex match (return all lines starting with "Stopped")
Get-Service | Out-String -Stream | Select-String "^Stopped"

# Case-sensitive regex match (will return nothing as should be "^Stopped")
Get-Service | Out-String -Stream | Select-String "^STOPPED" -CaseSensitive

# Case-sensitive non-regex match (-SimpleMatch forces a simple string match)
Get-Service | Out-String -Stream | Select-String "Stop" -CaseSensitive -SimpleMatch

# Get-Childitem "C:\Windows\" -Recurse -Include *.log -ErrorAction SilentlyContinue | Select-String "Error" -ErrorAction SilentlyContinue | Group-Object filename | Sort-Object Count -Descending
# ls "C:\Windows\" -i *.log -r -EA Silent | sls "Error" -EA Silent | group filename | sort Count -Descending
    
'@
    
echo $out
}




function Help-PowerShellLanguageReference {
    # https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/downloading-powershell-language-reference-or-any-file
    $url = "https://download.microsoft.com/download/4/3/1/43113f44-548b-4dea-b471-0c2c8578fbf8/powershell_langref_v4.pdf"
    # get desktop path
    $desktop = [Environment]::GetFolderPath('Desktop')
    $destination = "$desktop\langref.pdf"
    # enable TLS1.2 for HTTPS connections
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    # download PDF file
    Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
    # open downloaded file in associated program
    Invoke-Item -Path $destination
    # ToDo: Wait for the process started by Invoke-Item to exit, then delete the downloaded $destination file 
}

function Help-PowerShellMagazineQuick {
    # https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/downloading-powershell-language-reference-or-any-file
    $url = "https://download.microsoft.com/download/4/3/1/43113f44-548b-4dea-b471-0c2c8578fbf8/powershell_langref_v4.pdf"
    # get desktop path
    $desktop = [Environment]::GetFolderPath('Desktop')
    $destination = "$desktop\langref.pdf"
    # enable TLS1.2 for HTTPS connections
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    # download PDF file
    Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
    # open downloaded file in associated program
    Invoke-Item -Path $destination
}

function Help-Gci-FilterIncludeExclude {}   # This is just for all shitty issues around Include / Exclude
# Get-Command | Where { $_.parameters.keys -Contains "Filter" -And $_.Verb -Match "Get"}
# Get-Command | Where { $_.parameters.keys -Contains "Include" -And $_.Verb -Match "Get"}
# Get-Command | Where { $_.parameters.keys -Contains "Exclude" -And $_.Verb -Match "Get"}
# Get-Command | Where { $_.parameters.keys -Contains "Filter"}
# Get-Command | Where { $_.parameters.keys -Contains "Include"}
# Get-Command | Where { $_.parameters.keys -Contains "Exclude"}
# ONE DISADVANTAGE OF -FILTER
# You can only sift using one value with -Filter, whereas  -Include can accept multiple values, for example ".dll, *.exe"
# https://www.computerperformance.co.uk/powershell/file-gci-filter/

# Good explanation of the mess by Alex K. Angelopoulos from 2013 here: https://social.technet.microsoft.com/Forums/en-US/62c85fc4-1d44-4c3a-82ea-d49109423471/inconsistency-between-getchilditem-include-and-exclude-parameters?forum=winserverpowershell
# A more up to date explanation from mklement0 (2016, but updated May 2022): https://stackoverflow.com/questions/38269209/using-get-childitem-exclude-or-include-returns-nothing/38308796#38308796
# The mklement0 answer explains how some of the issues have been fixed in PS 7.x but will never be fixed in 5.x
# 
### -Filter is the most useful parameter to refine output of PowerShell cmdlets (much better than -Include or -Exclude).
# Get-ChildItem $Env:windir\System32 -Filter *.dll
# gci $Env:windir\System32\*.dll   # This works just as well
# # Always choose -Filter rather than -Include.  Filtering is faster, and the results are more predictable.
# [Decimal]$Timing = ( Measure-Command {
#     $Result = Get-ChildItem $env:windir\System32 -Filter *.dll -Recurse -EA SilentlyContinue
# } ).TotalMilliseconds
# $Timing =[math]::round($Timing)
# Write-Host "TotalMilliseconds" $Timing "`nFile count:" $Result.count
# 
# $Timing = (Measure-Command {
#     $Result = Get-ChildItem $env:windir\System32 -Include *.dll -Recurse -EA SilentlyContinue
# } ).TotalMilliseconds
# $Timing =[math]::round($Timing)
# Write-Host "TotalMilliseconds" $Timing "`nFile count:" $Result.count

# Get-ChildItem had a complex history with compromises between usability and standards compliance.
# The original design was as a tool for enumerating items in a namespace; similar to but not equivalent to dir and ls.
# The syntax and usage was going to conform to standard PowerShell (Monad at the time) guidelines. i.e. the Path
# parameter would have truly just meant Path; it would not have been usable as a combination path specification and
# result filter, which is what it is now.
# dir c:\temp     # return children of the container c:\temp.
# dir c:\temp\*   # return children of all containers inside c:\temp. With (2), you would never get c:\tmp\a.txt returned, since a.txt is not a container.
# There are reasons that this was a good idea. The parameter names and filtering behavior was consistent with the evolving PowerShell design standards, and best of all the tool would be straightforward to stub in for use by namespace providers consistently.
# However, this produced a lot of heated discussion. A rational, orthogonal tool would not allow the convenience we get with the dir command for doing things like this:
# dir c:\tmp\a*.txt  # Possibly more important was the "crash" factor.  It's so instinctive for admins to do things like (3) that our fingers do the typing when we list directories, and the instant failure or worse, weird, dissonant output we would get with a more pure Path parameter is exactly like slamming into a brick wall.

# From the documentation: -Include <string[]> Retrieves only the specified items. The value of this
# parameter qualifies the Path parameter. Enter a path element or pattern, such as "*.txt". Wildcards
# are permitted. The Include parameter is effective only when the command includes the Recurse parameter
# or the path leads to the contents of a directory, such as C:\Windows\*, where the wildcard character
# specifies the contents of the C:\Windows directory.
# Note: only effective with -Recurse, or when C:\Path\* leads to the contexts of a directory (!!)

# Get-File

function Help-Dir {
    $out = @'

Note: Try to not use 'ls' alias as it conflicts with PowerShell on Linux, Get-ChildItem / gci / dir  are fine
Mode (Attributes) column: l (link), d (directory), a (archive), r (read-only), h (hidden), s (system).
# gci C:\0\*.txt -Directory [-ad] -Hidden [-ah] -ReadOnly [-ar] -System [-as]
# gci C:\0\*.txt -adhrs -Recurse

### General:
Mode (Attributes) column: l (link), d (directory), a (archive), r (read-only), h (hidden), s (system).
gci C:\ -Name                    # Equivalent DOS:   dir C:\ /b
gci C:\0\*.txt -Recurse -Force   # -Force will also show hidden files etc
gci C:\0\*.txt -Recurse -Force -Include A*
gci C:\0\*.txt -Recurse -Force -Exclude A*
gci C:\0\ -Depth 1          # Works
gci C:\0\*.txt -Depth 1     # Broken, not sure why
gci C:\0\S* -Depth 2 -Dir   # Just show directories starting with S to a depth of 2
gci C:\0\*.txt -Recurse -Force -Filter "" update this ...

dir C:\0\*.txt -Attribute Archive, Compressed, Device, Directory, Encrypted, Hidden, IntegrityStream, Normal, `
NoScrubData, NotContentIndexed, Offline, ReadOnly, ReparsePoint, SparseFile, System, Temporary
gci C:\0\*.txt -Directory -Hidden -ReadOnly -System   # -ad -ah -ar -as
gci C:\0\*.txt -adhrs -Recurse

https://stackoverflow.com/questions/19091750/how-to-search-for-a-folder-with-powershell
https://stackoverflow.com/questions/55029472/list-folders-at-or-below-a-given-depth-in-powershell
gci C:\pr* *wind* -Recurse -Directory   

### Basic Dir with selection of items
Get-ChildItem C:\ | Where-Object { $_.Name -Like '*pr*' }
gci C:\ | ? Name -Like '*pr*'    # Demonstrates PSv3 shorthand instead of "{ $_.Name -Like 'xxx' }"
gci C:\*pr*
gci C:\*pr* | ForEach-Object { Remove-Item -LiteralPath $_.Name }   # Use the full literal path to perform the deletion
gci C:\*pr* | % { rm -Literal $_.Name }                             # Shorthand

### Dir commands with -filter and -include
A caveat: this command actually gets files like *.txt* (-Filter uses CMD wildcards). If this is not what you want then use -Include *.txt
https://stackoverflow.com/questions/13126175/get-full-path-of-the-files-in-powershell/13126266

### Recurse, find all items >500 MB in a given folder
Get-ChildItem C:\Windows -Recurse | Where-Pbject { $_.Length -gt 524288000 } | Sort-Object Length | Format-Table FullName,Length -Auto
dir C:\Win* -r | ? length -gt 500mb | sort length | ft fullname,length

### DOS equivalent to the above(!)
forfiles /P C:\ /M *.* /S /C "CMD /C if @fsize gtr 524288000 echo @PATH @FSIZE"
/P Path to process (if missing, start in current dir), /M (Mask) Search filter, /S (Sub-Directories), /C (Command),
@PATH is a variable. Shows the full path of the file. @FSIZE is a variable. Shows the file size, in bytes.

### Find all items with name like 'Win*' and show the name and LastWriteTime
Get-ChildItem -Path c:\Windows | Where-Object {$_ -like 'Win*'} | ForEach-Object { Write-Output -InputObject "$($_.Name) : $($_.LastWriteTime)" }
ls c:\windows | ? {$_ -like 'Win*'} | % {echo "$_.Name : $_.LastWriteTime"}

### Rename all file extensions in a folder:
Get-ChildItem *.jpeg | Rename-Item -newname { $_.name -replace '.jpeg','.jpg' }
dir -recurse *.jpeg | Rename-Item -newname { $_.name -replace '.jpeg','.jpg' }   # recursively
forfiles /S /M *.jpeg /C "cmd /c rename @file @fname.jpg"

### -Include: To search for multiple extensions (or any path pattern)
# Just need to make sure to add the '*' wildcard to the end of the path (it leads to the contents of the directory).
# But note that you can omit the '*' wildcard if you also specify -Recurse(!)
Get-Childitem -Path C:\* -include *.log,*.txt,*.nfo
dir C:\* -i *.log,*.txt,*.nfo

### -Include and -Exclude are unintuitive and contain a bug in PowerShell 5.1 that will never be fixed
https://stackoverflow.com/posts/38308796/revisions
https://syntaxfix.com/question/18836/how-can-i-exclude-multiple-folders-using-get-childitem-exclude
https://stackoverflow.com/questions/51666987/in-powershell-get-childitem-exclude-is-not-working-with-recurce-parameter
It is better to avoid -Exclude complately and instead use the following to exclude items:
gci -r -dir | ? fullname -notmatch 'dir1|dir2|dir3'   # Will exlucde folders dir1,dir2,dir3
But note that this will find all folders first, then exclude the ones not to match, so is not great performance-wise
(gci -recurse *.aspx,*.ascx).fullname -notmatch '\\obj\\|\\bin\\'   # \obj\ or \bin\   # remember that PowerShell is case-insensitive by default, so -inotmatch is not needed
gci -path c:\ -filter temp.* -exclude temp.xml   # Returns no results
gci -path c:\* -filter temp.* -exclude temp.xml  # Returns  all the temp.* files, except temp.xml, note the c:\* which causes this
gci $source -Directory -recurse | ? {$_.fullname -NotMatch "\\\s*_"} | % { $_.fullname }   # "\\\s*_" is regex for \, any amount of whitespace, _ 

### DeDup, using 'Group-Object'
$env:PSModulePath.Split(";") | gci -Directory | group Name | where Count -gt 1 | select Count,Name,@{ n = "ModulePath"; e = { $_.Group.Parent.FullName } }
"g:\TV\*", "h:\TV\*" | gci -i *.avi,*.mkv -Recurse | group Name | where Count -gt 1 | select Count,Name,@{ n = "Paths"; e = { $_.Group.Directory } }

'@
echo $out
}

function Help-Terminal-Icons {
    $out = @'

https://gist.github.com/markwragg/6301bfcd56ce86c3de2bd7e2f09a8839
How to get @DevBlackOps Terminal-Icons module working in PowerShell on Windows
Note: since version 0.1.1 of the module this now works in Windows PowerShell or PowerShell Core.

Download and install this version of Literation Mono Nerd Font which has been specifically fixed to be recognised as monospace on Windows:
https://github.com/haasosaurus/nerd-fonts/blob/regen-mono-font-fix/patched-fonts/LiberationMono/complete/Literation%20Mono%20Nerd%20Font%20Complete%20Mono%20Windows%20Compatible.ttf

(see this issue for more info: https://github.com/ryanoasis/nerd-fonts/issues/269)

Modify the registry to add this to the list of fonts for terminal apps (cmd, powershell etc.):
$key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Console\TrueTypeFont'
Set-ItemProperty -Path $key -Name '000' -Value 'LiberationMono NF'
Open PowerShell, right click the title bar > properties > Font > select your new font from the list.

Install and load Terminal-Icons:

Install-Module Terminal-Icons -Scope CurrentUser
Import-Module Terminal-Icons

# ====== Full code to automate configuration =====

### Install Terminal-Icons (get LiterationMono NF Nerd font, install, add required Concole registry key, then install Modules)
$url = 'https://github.com/haasosaurus/nerd-fonts/raw/regen-mono-font-fix/patched-fonts/LiberationMono/complete/Literation%20Mono%20Nerd%20Font%20Complete%20Mono%20Windows%20Compatible.ttf'
$name = "LiterationMono NF"
$file = "$($env:TEMP)\$($name).ttf"
Start-BitsTransfer -Source $url -Destination $file   # Download the font

$Install = $true  # $false to uninstall (or 1 / 0)
$FontsFolder = (New-Object -ComObject Shell.Application).Namespace(0x14)   # Must use Namespace part or will not install properly
$filename = (Get-ChildItem $file).Name
$filepath = (Get-ChildItem $file).FullName
$target = "C:\Windows\Fonts\$($filename)"

If (Test-Path $target -PathType Any) { Remove-Item $target -Recurse -Force } # UnInstall Font
# Following action performs the install, requires user to click on yes
If ((-not(Test-Path $target -PathType Container)) -and ($Install -eq $true)) { $FontsFolder.CopyHere($filepath, 16) }

# Need to set this for console
$key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Console\TrueTypeFont'
Set-ItemProperty -Path $key -Name '000' -Value $name

# Following are all required to enable the Terminal-Icons 'DevBlackOps' Theme
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force   # Always need this, required for all Modules
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted   # Set Microsoft PowerShell Gallery to 'Trusted'
Install-Module Terminal-Icons -Scope CurrentUser
Import-Module Terminal-Icons
Install-Module WindowsConsoleFonts
Set-ConsoleFont $name
Set-TerminalIconsColorTheme -Name DevBlackOps   # After the above are setup, can add just this line to $Profile to always load DevBlackOps

# ===== End of configuration for Literation font =====

$name = "Meslo LG M Regular"   #  Nerd Font Completely Mono Windows Compatible
$file = "$($env:TEMP)\$($name).ttf"
$Install = $true  # $false to uninstall (or 1 / 0)
$FontsFolder = (New-Object -ComObject Shell.Application).Namespace(0x14)   # Must use Namespace part or will not install properly
$filename = (Get-ChildItem $file).Name
$filepath = (Get-ChildItem $file).FullName
$target = "C:\Windows\Fonts\$($filename)"
# If (Test-Path $target -PathType Any) { Remove-Item $target -Recurse -Force } # UnInstall Font
If ((-not(Test-Path $target -PathType Container)) -and ($Install -eq $true)) { $FontsFolder.CopyHere($filepath, 16) }
$key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Console\TrueTypeFont'
Set-ItemProperty -Path $key -Name '000' -Value $name
# Following are all required to enable the Terminal-Icons 'DevBlackOps' Theme
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force   # Always need this, required for all Modules
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted   # Set Microsoft PowerShell Gallery to 'Trusted'
Install-Module Terminal-Icons -Scope CurrentUser
Import-Module Terminal-Icons
Install-Module WindowsConsoleFonts
Set-ConsoleFont $name
Set-TerminalIconsColorTheme -Name DevBlackOps   # After the above are setup, can add just this line to $Profile to always load DevBlackOps

Meslo LG M Regular Nerd Font Completely Mono Windows Compatible
'@

echo $out
}


function Help-EbooksCalibre {
    $out = @'

Downloading from Gutenberg, files are named like "pg82.mobi", but they contain the meta-data
so can rename the files correctly using the Calibre ebook-meta tool to get the details.
Need to remember the "Out-String -Stream" trick to make sure output are per line(!)

foreach ($i in dir) {
    $all = (ebook-meta $i)
    $title = ($all | sls "^Title" | Out-String -Stream).replace("Title               : ", "").trim(" ")
    $author = ($all | sls "^Author" | Out-String -Stream).replace("Author              : ", "").replace("Author(s)           : ", "").trim(" ")
    $out = ("$title - $author").replace(":", ";").replace("   ", " ").replace("  ", " ").replace("^ ", "").trimstart(" ").trimend(" ")
    # $ext = ($i | select Extension).Extension   # Also select Basename
    echo "$($out).mobi"   # Can't get this to work, just returns ".@{Extension= ..mobi}"
    mv $i "$($out).mobi"
}



'@

echo $out
}


function Help-RegEx {
    $out = @'

Collect some useful regex that I've used in here.

gc $profile | sls -Pattern '^if'              # Show all lines that start with 'if' as a 'MatchInfo' object
gc $profile | sls -Pattern '^if' -NotMatch    # Show all lines that do not start with 'if'
gc $profile | sls -Pattern 'and' -AllMatches  # Repeatedly look for 'and' (by default, will only match the first hit in a line)
$Events = Get-EventLog -LogName application -Newest 100
$Events | Select-String -InputObject {$_.message} -Pattern "failed"

https://powershellexplained.com/2017-07-31-Powershell-regex-regular-expression/
http://www.thejoyofcode.com/Powershell_to_test_your_Regexes.aspx
https://www.regexbuddy.com/powershell.html

'@
echo $out
}

function Help-WIP {
    $out = @'

Essentials to know!!!
 
SEARCH COMMAND HISTORY:
Press Ctrl+r, a new subprompt will appear "bck-i-search:".
Now enter some text, the newest match to that text will be shown.
Continue to press Ctrl+r to go backward (older) until you find the command required.
To move foward (newer) in the command history, use Ctrl+s "fwd-i-search:".
<get some notes on historypx here?>

PSReadLine HELP:
Ctrl+Alt+Shift+?

[System.Environment]::GetEnvironmentVariables()   # Show all Environment Variables
gci env:*     *or*      gci env:                  # Also show all Environment Variables


Random tips:

(123.456).ToString("C")       -> £123.46     # Currency, note, in PS, just takes locale, you cannot do .ToString("C", fr-Fr) to get other 
(5/21).ToString("P")          -> 23.81%      # Percentage
(-1052.032911).ToString("e8") -> -1.052e+003 # Exponential (Scientific)
1234 ("D") -> 1234 , -1234 ("D6") -> -001234 # Decimal, will retain negative sign if required
255 ("X") -> FF, 255 ("X8") -> 000000FF      # Hex
More here: https://docs.microsoft.com/en-us/dotnet/standard/base-types/standard-numeric-format-strings

To time a command:
powershell -noprofile -ExecutionPolicy Bypass ( Measure-Command { powershell "Write-Host 1" } ).TotalSeconds
( Measure-Command { <command to run> } ).TotalSeconds

$now = Get-Date -format "yyyy-MM-dd__HH-mm-ss"    # filename-compatible datetime format that I use for logs etc
Get-WindowsCapability -Online | Where-Object { $_.State -eq 'Installed' }   # Show installed DISM components

# Convert a Here-String to an array in one step with multiple carriage return types:
$HereStringSample=@'
Banana
Raspberry
`'@
$HereStringSample.Split(@("$([char][byte]10)", "$([char][byte]10)", "$([char][byte]13)", [StringSplitOptions]::None))

# Show 'platform' (this only works on PS Core versions)
$PSVersionTable.Platform

Should build notes on EventLog stuff, always useful ...

'@
echo $out | more

}

function Get-OpenWindows { Get-Process | where {$_.mainWindowTitle} | Format-Table id, name, mainwindowtitle -autosize }   # Just show all open windows

function Install-TrustedSubnet ($ThirdOctet) {   # Quickly set local subnet as TrustedHosts.
    if ($ThirdOctet -eq $null) { "nothing entered..." } 
    else {
        "Current TrustedHosts (Get-Item WSMan:\localhost\Client\TrustedHosts):"
        Get-Item WSMan:\localhost\Client\TrustedHosts
        ""
        "Press any key to trust '192.168.$ThirdOctet.*'"
        "Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""192.168.$ThirdOctet.*"" -Force"
        pause
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.$ThirdOctet.*" -Force
        Get-Item WSMan:\localhost\Client\TrustedHosts
        ""
    }
}


function Enable-WinRM {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host 'Enable WinRM and add Trusted subnets (192.168.0.* or 192.168.1.*)' -ForegroundColor Green
    Write-Host ""
    Write-Host "WinRM is mostly used with Active Dirctory but I will mostly use it in a"
    Write-Host "WORKGROUP which requires a few different configurations like TrustedHosts"
    Write-Host ""
    Write-Host "Enable-PSRemoting -Force -SkipNetworkProfileCheck"
    Write-Host 'winrm enumerate winrm/config/Listener'
    Write-Host 'Get-WSManInstance -ResourceURI winrm/config/listener -SelectorSet @{Address="*";Transport="http"}'
    Write-Host 'netsh advfirewall firewall add rule name="WinRM-HTTP"  dir=in localport=5985 protocol=TCP action=allow'   # HTTP
    Write-Host 'netsh advfirewall firewall add rule name="WinRM-HTTPS" dir=in localport=5986 protocol=TCP action=allow'   # HTTPS
    Write-Host ""
    Write-Host "After this setup completes, make sure to setup the subnet as Trusted"
    Write-Host "   Register-TrustedSubnet <Third-Octect>"
    Write-Host "Replace <Third-Octet> by 0 to trust 192.168.0.* networks"
    Write-Host "Replace <Third-Octet> by 1 to trust 192.168.1.* networks"
    Write-Host "Without this, home WORKGROUP setups will not work with WinRM"
    Write-Host ""
    Write-Host "To connect to a remote host, the local hosts IP must be Trusted, then type:"
    Write-Host "    Enter-PSSession -ComputerName 192.168.0.21 -Credential (Get-Credential)"
    Write-Host ""
    Write-Host "On the remote system:"
    Write-Host '    Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.0.*" -Force'
    Write-Host "    Get-Item WSMan:\localhost\Client\TrustedHosts"
    Write-Host ""
    Write-Host "========================================`n" -ForegroundColor Green
    Write-Host ""

    if (! (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) ) {
        Write-Host "Must be Administrator to configure WinRM"
        Write-Host ""
        break
    }

    if ((Read-Host "Would you like to enable WinRM on this host (default is 'n') [y/n]? ") -ne "y" ) {
        Write-Host "WinRM configuration was cancelled"
        Write-Host ""
        break
    }

    # [HPEnvy] Connecting to remote server HPEnvy failed with the following error message : The WinRM client cannot process the request. If the authentication scheme is different
    # from Kerberos, or if the client computer is not joined to a domain, then HTTPS transport must be used or the destination machine must be added to the TrustedHosts
    # configuration setting. Use winrm.cmd to configure TrustedHosts. Note that computers in the TrustedHosts list might not be authenticated. You can get more information about
    # that by running the following command: winrm help config. For more information, see the about_Remote_Troubleshooting Help topic.
    #     + CategoryInfo          : OpenError: (HPEnvy:String) [], PSRemotingTransportException
    #     + FullyQualifiedErrorId : ServerNotTrusted,PSSessionStateBroken
    #
    # winrm.cmd
    #
    # Windows Remote Management Command Line Tool
    # 
    # Windows Remote Management (WinRM) is the Microsoft implementation of
    # the WS-Management protocol which provides a secure way to communicate
    # with local and remote computers using web services.
    # 
    # Usage:
    #   winrm OPERATION RESOURCE_URI [-SWITCH:VALUE [-SWITCH:VALUE] ...]
    #         [@{KEY=VALUE[;KEY=VALUE]...}]
    # 
    # For help on a specific operation:
    #   winrm g[et] -?        Retrieving management information.
    #   winrm s[et] -?        Modifying management information.
    #   winrm c[reate] -?     Creating new instances of management resources.
    #   winrm d[elete] -?     Remove an instance of a management resource.
    #   winrm e[numerate] -?  List all instances of a management resource.
    #   winrm i[nvoke] -?     Executes a method on a management resource.
    #   winrm id[entify] -?   Determines if a WS-Management implementation is
    #                         running on the remote machine.
    #   winrm quickconfig -?  Configures this machine to accept WS-Management
    #                         requests from other machines.
    #   winrm configSDDL -?   Modify an existing security descriptor for a URI.
    #   winrm helpmsg -?      Displays error message for the error code.
    # 
    # For help on related topics:
    #   winrm help uris       How to construct resource URIs.
    #   winrm help aliases    Abbreviations for URIs.
    #   winrm help config     Configuring WinRM client and service settings.
    #   winrm help certmapping Configuring client certificate access.
    #   winrm help remoting   How to access remote machines.
    #   winrm help auth       Providing credentials for remote access.
    #   winrm help input      Providing input to create, set, and invoke.
    #   winrm help switches   Other switches such as formatting, options, etc.
    #   winrm help proxy      Providing proxy information.

    # Attempt to setup main WinRM steps.
    # Must make sure that all networks are Private for this to work, or do I? Maybe not, just use skipnetworkprofilecheck
    # Set-WSManQuickConfig is always a disaster
    # Set Network location private for all networks:
    # https://devblogs.microsoft.com/powershell/setting-network-location-to-private/
    # https://winaero.com/blog/network-location-type-powershell-windows-10/
    # https://www.petri.com/powershell-remoting-tip-setting-a-network-category-to-private
    # https://www.jorgebernhardt.com/how-to-force-a-network-profile-in-windows-using-powershell/
    # https://www.tenforums.com/tutorials/6815-set-network-location-private-public-domain-windows-10-a.html?__cf_chl_jschl_tk__=301f3d0cab8cd80583db1ef23f88e1182066f27f-1585677722-0-Ae-1ogRWPkRiQKXDElWTmiQwhweekx5PJuej82kH-h3W7F2fuuhYV3BzefOeap1e0A0wwgcqpmeVvvG2MaTGq7iltikCZbEzleZSPoStNho5WWekaPMZFxKIGGSkW-q-pcyZGTtB4K2DLV0HmWmovhhYijQTUbeLWI_f93pC9-W96oT4bAo0Xz0pvOjwax1iXnXaRrOqFFA6qtdMlb7iWot6Kv8-YcAMKuKBz7DzP5450jLLHO9ZAovwgViMdNcHcWGi0cP-F_MRKgjXUMn2O2-h4I-JK4luqpViHpaj1C1XlwnFXQtoYYdpxAZaA3-zXwFJSZcw5eV-_sq3E82Ia_DxP9wn1_TNxwh2DZxFopbAM_ESTajO1je1YL-uPkbyX_T6u8K2MMZQ0pIjAljVqkYnBVc3cZmRyX2j3x6afbH3dhLBRyn6M4Wj6v7JpCeIRQ
    
    ####################
    # 1. Run Enable-PSRemoting
    ####################
    "Enable-PSRemoting -Force -SkipNetworkProfileCheck"
    "========================================`n"
    Enable-PSRemoting -Force -SkipNetworkProfileCheck
    # To avoid the error message and enable PowerShell remoting on a public network, you can use the -SkipNetworkProfileCheck parameter.
    # psexec.exe \\RemoteComputerName -s powershell Enable-PSRemoting -Force -SkipNetworkProfileCheck   # PS Tools should be installed on all systems!
    # Invoke-Command -VMName <VM_name> -ScriptBlock {Enable-PSRemoting -Force -SkipNetworkProfileCheck} -Credential Administrator
    # https://4sysops.com/wiki/enable-powershell-remoting/
    # winrm quickconfig   # Old way to enable
    # Set-WSManQuickConfig   # <-- This will fail if any Networks are public!
    # To view the current listeners that are running on the WinRM service, run the following command:
    "winrm enumerate winrm/config/Listener"
    "========================================`n"
    winrm enumerate winrm/config/Listener
    'Get-WSManInstance -ResourceURI winrm/config/listener -SelectorSet @{Address="*";Transport="http"}'
    "========================================`n"
    Get-WSManInstance -ResourceURI winrm/config/listener -SelectorSet @{Address="*";Transport="http"}

    ####################
    # 2. Open the firewall:
    ####################
    "Setup Firewall Rules: Open 5985 for HTTP and 5986 for HTTPS"
    "========================================`n"
    'netsh advfirewall firewall add rule name="WinRM-HTTP"  dir=in localport=5985 protocol=TCP action=allow   # HTTP'
    netsh advfirewall firewall add rule name="WinRM-HTTP"  dir=in localport=5985 protocol=TCP action=allow   # HTTP
    'netsh advfirewall firewall add rule name="WinRM-HTTPS" dir=in localport=5986 protocol=TCP action=allow   # HTTPS'
    netsh advfirewall firewall add rule name="WinRM-HTTPS" dir=in localport=5986 protocol=TCP action=allow   # HTTPS

    ####################
    # 3. If accessing via cross platform tools like chef, vagrant, packer, ruby or go, run these:
    ####################
    # Note: DO NOT use the above winrm settings on production nodes. This should be used for tets instances only for troubleshooting WinRM connectivity.
    # winrm set winrm/config/client/auth '@{Basic="true"}'
    # winrm set winrm/config/service/auth '@{Basic="true"}'
    # winrm set winrm/config/service '@{AllowUnencrypted="true"}'

    ####################
    # 4. Workgroups: In a workgroup environment, you have to add the IP addresses of all Trusted computers to the TrustedHosts list manually.
    ####################
    $DefaultIPAddress = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].IPAddress[0]
    $DefaultThirdOctet = ($DefaultIPAddress -split "\.")[2]   # Third Octet is [2] as count from 0 !
    "Default IP Address: $DefaultIpAddress"
    "Third Octect (to define subnet): $DefaultThirdOctet"
    if ($DefaultThirdOctect -eq 0) { Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.0.*" -Force }
    if ($DefaultThirdOctect -eq 1) { Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.*" -Force }
    "Get-Item WSMan:\localhost\Client\TrustedHosts   # View TrustedHosts values"
    "========================================`n"
    Get-Item WSMan:\localhost\Client\TrustedHosts

    ####################
    # Enter-PSSession -ComputerName 192.168.0.21 -Credential (Get-Credential)
    ####################

    # $current=(get-item WSMan:\localhost\Client\TrustedHosts).value   # good to query first, then add
    # PS C:\> $current+=",testdsk23,alpha123"
    # PS C:\> set-item WSMan:\localhost\Client\TrustedHosts -value $current
    # Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force   # dangerous

    # Enable-PSRemoting will perform the following steps:
    #    Runs the Set-WSManQuickConfig cmdlet, which performs the following tasks:
    #    Starts the WinRM service.
    #    Sets the startup type on the WinRM service to Automatic.
    #    Creates a listener to accept requests on any IP address.
    #    Enables a firewall exception for WS-Management communications.
    #    Registers the Microsoft.PowerShell and Microsoft.PowerShell.Workflow session configurations, if it they are not already registered.
    #    Registers the Microsoft.PowerShell32 session configuration on 64-bit computers, if it is not already registered.
    #    Enables all session configurations.
    #    Changes the security descriptor of all session configurations to allow remote access.
    #    Restarts the WinRM service to make the preceding changes effective.

    # Troubleshoot connectivity issues between Win 10 / Win 7
    # https://superuser.com/questions/1351882/windows-10-cannot-connect-to-windows-7-computers
    # https://stackoverflow.com/questions/21548566/how-to-add-more-than-one-machine-to-the-trusted-hosts-list-using-winrm
    # https://winintro.ru/windowspowershell2corehelp.en/html/f23b65e2-c608-485d-95f5-a8c20e00f1fc.htm
    # http://www.hurryupandwait.io/blog/understanding-and-troubleshooting-winrm-connection-and-authentication-a-thrill-seekers-guide-to-adventure
    # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/enable-psremoting?view=powershell-5.1
    # https://www.visualstudiogeeks.com/devops/how-to-configure-winrm-for-https-manually
    # https://www.dtonias.com/add-computers-trustedhosts-list-powershell/
    # https://www.poftut.com/enable-powershell-remoting-psremoting-winrm/

    # Module to make dealing with trusted hosts slightly easier, psTrustedHosts.
    # Add-TrustedHost, Clear-TrustedHost, Get-TrustedHost, and Remove-TrustedHost.
    #    Install-Module psTrustedHosts -Force

    # Test connectivity to PS Sessions:
    # echo 'Enter-PSSession -ComputerName IP -Credential (Get-Credential -UserName administrator -Message "Give me the password please")'
    # Test-WSMan -ComputerName <IP or host name>

    # By default two BUILTIN groups are allowed to use PowerShell Remoting as of v4.0. The Administrators and Remote Management Users.
    # (Get-PSSessionConfiguration -Name Microsoft.PowerShell).Permission
    # Sessions are launched by Secret Server under the user's context which means all the same security controls and policies apply within the session.
    # Investigating PowerShell Attacks by FireEye
    # Your environment may already be configured for WinRM. If your server is already configured for WinRM but is not using the default configuration, you can change the URI to use a custom port or URLPrefix.
    # Configuration (Standalone)
    # By default WinRM uses Kerberos for Authentication. Since Kerberos is not available on machines which are not joined to the domain - HTTPS is required for secured transport of the password. Only use this method if you are going to be running scripts from a Secret Server Web Server or Distributed Engine which is not joined to the domain.
    # Note: WinRM HTTPS requires a local computer "Server Authentication" certificate with a CN matching the hostname, that is not expired, revoked, or self-signed to be installed. A certificate would need to be installed on each endpoint for which Secret Server or the Engine would manage.
    # Create the new listener:
    # New-WSManInstance - ResourceURI winrm/config/Listener -SelectorSet @{Transport=HTTPS} -ValueSet @{Hostname="HOST";CertificateThumbprint="XXXXXXXXXX"}
}

function Cleanup-DISM {   # DISM (Deployment Image Servicing and Management tool) to cleanup Service Packs and components.
    # ToDo: update this to include command line CleanUp tasks
    "
Script will run all DISM (Deployment Image Servicing and Management tool) cleanup options
including removal of old Service Packs etc:

    dism /Online /Cleanup-Image /StartComponentCleanup
    dism /Online /Cleanup-Image /StartComponentCleanup /ResetBase
    dism /Online /Cleanup-Image /SPsuperseded`n"

    function Test-Administrator {
        $user = [Security.Principal.WindowsIdentity]::GetCurrent();
        (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    if (Test-Administrator -eq $true) {
        $rundism = read-host "Run 'DISM' cleanup actions (default is 'n') [y/n]? "
        if ($rundism -eq 'y') {
            dism /Online /Cleanup-Image /StartComponentCleanup
            dism /Online /Cleanup-Image /StartComponentCleanup /ResetBase
            dism /Online /Cleanup-Image /SPsuperseded
        }
    }
    else {
        "Current user does not have elevated Administrator privileges, so script will exit."
    }
    ""
}


####################
#
# Random Notes
#
####################

# Ping notes... Ping-IPRange is faster than any other method, installed with sample scripts, but these were useful
# https://www.thomasmaurer.ch/2016/02/basic-networking-powershell-cmdlets-cheatsheet-to-replace-netsh-ipconfig-nslookup-and-more/
# https://www.thomasmaurer.ch/2013/10/superping-powershell-test-netconnection/
# https://www.thomasmaurer.ch/2019/09/how-to-enable-ping-icmp-echo-on-an-azure-vm/
# https://www.petri.com/building-ping-sweep-tool-powershell

# Retry this, but can use with extra that installs the font:
# if (Get-Command Set-TerminalIconsColorTheme -EA silent) { Set-TerminalIconsColorTheme -Name DevBlackOps }
# jobs, like so : $MyJob = Start-Job -ScriptBlock {$MyCommand}
# And when the job has completed, you can get its output, like so : Receive-Job -Job $MyJob
# Could turn on transcript for all sessions ? Same them in the $profile folder, same name as Profile with _yyyy-mm-dd__hh_mm.txt on them?



# Something to work on ... ways to update the DOS prompt somehow ...
# PROMPT=$E[33m$D$S$T$H$H$H$S$E[37m$M$_$E[1m$P$G
# $e[xxm -- is a color setting $e[37m is normal white
# $d     -- current date
# $t     -- current time
# $h     -- backspace (to remove the milliseconds from $t)
# $p     -- current path/directory
# $g     -- ">" character
# $s     -- space character
# Also this ...
# https://www.windowscentral.com/how-change-appearance-command-prompt-windows-10
# https://www.tenforums.com/tutorials/94089-change-screen-buffer-size-console-window-windows.html



# https://community.spiceworks.com/topic/664020-maximize-an-open-window-with-powershell-win7
# function Set-WindowStyle {
#     Write-Verbose ("Set Window Style {1} on handle {0}" -f $MainWindowHandle, $($WindowStates[$style]))
#
# Very annoying when Windows positions console off bottom of screen etc, this will maximise all PowerShell instances
# function psmax { (Get-Process -Name powershell).MainWindowHandle | foreach { Set-WindowStyle MAXIMIZE $_ } }

# [Edit:] Just a gotcha/caveat though: by default, outputs from PowerShell will be truncated on screen. This is both sensible (because outputs are often representations of objects) and frustrating (because Path in env:* is then truncated and you lose info). You just need to be aware of this, and remember ft -wrap -autosize and ConvertTo-Csv:
# ft -wrap -autosize can be ok but has a problem that Path will then split to multiple lines (not ideal for further manipulation)
# gci env:* | ft -wrap -autosize | out-string -stream | select-string "Pro"
# So the better fix is ConvertTo-Csv which will guarantee that all of the information will be retained and correctly held to the line that it should be in. As an advantage, it implicitly converts to string in doing this so you don't need out-string -stream anymore:
# gci env:* | ConvertTo-Csv | sls "Pro"
# But it outputs all Properties (see gci env:* | Get-Member) so it's best to just select (Select-Object) the properties that that you need first:
# gci env:* | select Name,Value | ConvertTo-Csv | sls "Pro"
# Quick-Grep ... techiniques like grep

# [Regex]::Matches($_, "^function ([a-z.-]+)","IgnoreCase").Groups[1].Value
# } | Where-Object { $_ -ine "prompt" } | Sort-Object
# Write-Host "Functions in profile extensions:" -F Yellow
# Get-Content -Path "$($profile)_extensions.ps1" | Select-String -Pattern "^function.+" | ForEach-Object {
#     [Regex]::Matches($_, "^function ([a-z.-]+)","IgnoreCase").Groups[1].Value
# } | Where-Object { $_ -ine "prompt" } | Sort-Object

# Should maybe backup Path at times?
# HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment
# HKEY_CURRENT_USER\Environment
# Determine all .NET installed on system and add latest to Path
# function mod($module) {
#     Get-Module -All
#     Get-Command -Module $module
#     $x = ((hs get-module | Out-String).replace("Get-Module ", "") -split("\r\n") -match '\S')
# }

# https://gist.github.com/stuartleeks/aaa4a55ebe4df55166cc
# Find-InPath notepad*
# Find-InPath notepad.exe
function Find-InPath {
    [CmdletBinding()] param ( [string] $filename )

    $matches = $env:Path.Split(';') | % { Join-Path $_ $filename} | ? { Test-Path $_ }
    if ($matches.Length -eq 0) { "No matches found" }
    else { $matches }
}

# Path "C:\xxx" -Add / -Remove , or Add-Path / Remove-Path
# just make this a bloody function as I type it by instinct a lot!
# Path -System  (just show system Path!)
# Path -User  (just show user Path!)
# Path -AddPath (system) -RemovePath (system)
# Path -AddPathUser (system) -RemovePathUser (system)
# Also have to add to current path after updating the registry! Could just do $env:Path = $registry version after update
# https://jkeohan.wordpress.com/2012/02/10/powershell-adding-directories-to-path-statement/
# https://devblogs.microsoft.com/scripting/use-powershell-to-modify-your-environmental-path/
# Need to split in case a matched path is a parent folder of the path to add
# Only adds path if not already found on path.
Function AddTo-SystemPath {
    Param( [array]$PathToAdd )
    $VerifiedPathsToAdd = $Null
    $PathArray = $Env:path -Split ";" -replace "\\+$", ""
    foreach ($Path in ($PathToAdd | % { $_.TrimEnd('\') })) {
        if ($PathArray -contains $Path ) {
            Write-Host "Current item in path is: $Path"
            Write-Host "$Path already exists in Path statement"
        } else {
            $VerifiedPathsToAdd += ";$Path"
            Write-Host "`$VerifiedPathsToAdd updated to contain: $Path"
        }
        if($VerifiedPathsToAdd -ne $null) {
            Write-Host "`$VerifiedPathsToAdd contains: $verifiedPathsToAdd"
            Write-Host "Adding $Path to Path statement now ..."
            [Environment]::SetEnvironmentVariable("Path", $env:Path + $VerifiedPathsToAdd, "Process")
        }
    }
}

# Essential paths to add:
# AddTo-SystemPath "C:\ProgramData\Scripts" | Out-Null
# Need to trum trailing "\"
# Need to uppercase start of string if c:\ or d:\ etc
# Need to implement switches -Add / -Remove (default is -Add)
# Need to also put added path into currently loaded path
# Find paths duplicated between User and Machine entries in reg
# Modules function to do something similar
# https://powershell.org/forums/topic/understanding-switch-parameters/
function Path ($AddPath, $RemovePath, [switch]$System, [switch]$User, [switch]$Loaded, [switch]$Quick) {
    <#
    .SYNOPSIS
    Shows the current path (loaded in session) and System path (from registry) and User path (from registry).
    Type 'Path' on its own to display all paths (currently loaded path + system and user paths)

    Path -System [-Quick]   Display the System path, split and sorted. Add -q to show the raw registry value
    Path -User [-Quick]     Display the User path, split and sorted. Add -q to show the raw registry value
    Path -Loaded [-Quick]   Display the currently loaded path in this session, split and sorted. Add -q to show the raw registry value

    Path -AddPath <path>    Add path to the System path (as this is default switch, can just use 'Path <path>')
    Path -RemovePath <path> Remove path from the System path    
    .EXAMPLE
    .
    Path -AddPath "C:\Program Files\Calibre2"      # -AddPath is default so can be omitted
    Path -RemovePath "C:\Program Files\Calibre2"   # -RemovePath is always required as it is not the default switch
    #>

    # Note: use the "." on a line by itself for .EXAMPLE to space out multiple examples (otherwise get crammed together)
    # Discussion on short pathnames: https://stackoverflow.com/questions/31547104/how-to-get-the-value-of-the-path-environment-variable-without-expanding-tokens

    # [Environment]::GetFolderPath('MyDocuments')
    # User Path:   (Get-Item -path "HKCU:\Environment" ).GetValue('Path', '', 'DoNotExpandEnvironmentNames')
    # System Path: (Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" ).GetValue('Path', '', 'DoNotExpandEnvironmentNames')
    $PathFull = [Environment]::GetEnvironmentVariable("Path")   # HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment
    $PathFullArray = $PathSystem -Split ";" -replace "\\+$", "" -replace "^;", "" -replace ";$", "" | ? { $_ -ne ""} | sort   # The ? (where) eliminates blank entries, leave the paths unsorted
    
    $PathSystem = [Environment]::GetEnvironmentVariable("Path", "Machine")   # HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment
    $PathSystemArray = $PathSystem -Split ";" -replace "\\+$", "" -replace "^;", "" -replace ";$", "" | ? { $_ -ne ""} | sort   # The ? (where) eliminates blank entries, leave the paths unsorted
    $regSystemEnv = "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment"   # Note the ":". Without this, PSProvider path is not valid
    $regSystemEnvPath = (Get-ItemProperty $regSystemEnv -Name Path).Path.TrimEnd(";")

    $PathUser = [Environment]::GetEnvironmentVariable("Path", "User")   # HKEY_CURRENT_USER\Environment
    $PathUserArray = $PathUser -Split ";" -replace "\\+$", "" -replace "^;", "" -replace ";$", "" | ? { $_ -ne ""} | sort   # The ? (where) eliminates blank entries, leave the paths unsorted
    $regUserEnv = "HKCU:\Environment"   # Note that the ":" is very important to specify using a PSProvider
    $regUserEnvPath = (Get-ItemProperty $regUserEnv -Name Path).Path.TrimEnd(";")
    
    if ($AddPath -ne $null) {
        # First add it to the registry then add it to currently loaded path ...
        # But which to add to?? Add to System if running as Admin and User otherwise?
        # Test for the existence of the path to add in the system, then the user paths
        if ($PathSystemArray -contains $AddPath) { echo "'$AddPath' is already in the system path..." ; break }
        if ($PathUserArray -contains $AddPath) { echo "'$AddPath' is already in the user path..." ; break }

        # Test that the path exists, don't add non-existent paths, maybe offer to create the path if not there?
        if (! (Test-Path $AddPath)) { echo "'$AddPath' does not exist, so will not update system path..." ; break }
        # Update System path
        echo "$regSystemEnvPath;$AddPath"
        Set-ItemProperty $regSystemEnv -Name Path -Value "$regSystemEnvPath;$AddPath"   # Add the new path separated by ";"
        # Update currently loaded path

        # [Environment]::SetEnvironmentVariable( "Path", $($OutputList -join ';'), [System.EnvironmentVariableTarget]::Machine )
        # This changes the registry key HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment which requires elevation - DOES THIS WORK???

        break
    }

    $PathArray = $env:Path -Split ";" -replace "\\+$", "" | sort
    
    if ($System -eq $true) {
        $PathSystem = [Environment]::GetEnvironmentVariable("Path", "Machine")   # HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment
        ""
        if ($Quick -eq $true) { $PathSystem }
        else {
            $PathArray = $PathSystem -Split ";" -replace "\\+$", "" -replace "^;", "" -replace ";$", "" | ?{$_} | sort
            # The ?{$_} used to be ($_ -ne "") but don't need '-ne ""'! This eliminates empty path entries ";;"
            # Note that $PathSystem contains warts and all, but $PathArray is the cleaned set of paths
            # $PathArray.Split('', [System.StringSplitOptions]::RemoveEmptyEntries), RemoveEmptyEntries is another option
            foreach ($Path in $PathArray) { $Path = $Path.TrimEnd('\') ; echo $Path }  # % { $_.TrimEnd('\') })) { echo $Path }    
        }
        ""
        break
    }
    if ($User -eq $true) {
        $PathUser = [Environment]::GetEnvironmentVariable("Path", "User")   # HKEY_CURRENT_USER\Environment
        ""
        if ($Quick -eq $true) { $PathUser }
        else {
            $PathArray = $PathUser -Split ";" -replace "\\+$", "" -replace "^;", "" -replace ";$", "" | ? { $_ } | sort  
            foreach ($Path in $PathArray) { $Path = $Path.TrimEnd('\') ; echo $Path }
        } 
        ""
        break
    }
    if ($Loaded -eq $true) {
        $PathLoaded = [Environment]::GetEnvironmentVariable("Path")            # Just get currently loaded session
        ""
        if ($Quick -eq $true) { $PathLoaded }
        else {
            $PathArray = $PathLoaded -Split ";" -replace "\\+$", "" -replace "^;", "" -replace ";$", "" | ? { $_ } | sort
            foreach ($Path in $PathArray) { $Path = $Path.TrimEnd('\') ; echo $Path }
        } 
        ""
        break
    }

    ""
    ":: Currently Loaded Path`n=========="
    $env:Path
    ""
    ":: HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment`n=========="
    [Environment]::GetEnvironmentVariable("Path", "Machine")
    ""
    ":: HKEY_CURRENT_USER\Environment`n=========="
    [Environment]::GetEnvironmentVariable("Path", "User")
    ""
    ":: Sorted Paths`n=========="
    foreach ($Path in $PathArray) { $Path = $Path.TrimEnd('\') ; echo $Path }  # % { $_.TrimEnd('\') })) { echo $Path }
    ""
    # https://www.computerperformance.co.uk/powershell/env-path/
    # (ls path).value.split() | ? { $_ -like "C:\Windows" -or $_ -like "%systemroot%" }
    echo '[Environment]::GetEnvironmentVariable("Path")          # Show currently loaded path (i.e. combination of "Machine" + "User")'
    echo '[Environment]::GetEnvironmentVariable("Path", "User")  # Show "User" part only, change to "Machine" for System Path)'
    echo '[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\bin", "Machine")   # Add a path to "Machine" (can also use "User"), updates only current session'
    echo 'rundll32 sysdm.cpl,EditEnvironmentVariables   # Open Environment Variables dialogue'
    ""
    echo '$RegistrySystemPath = "Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment"   # System PATH' 
    echo '$RegistrySystemPath = "Registry::HKEY_CURRENT_USER\Environment"   # User PATH, also at "HKCU:\Environment"'
    echo '(Get-ItemProperty -Path $RegistrySystemPath -Name Path).Path'
    echo '$PathArray = (Get-ItemProperty -Path $RegistrySystemPath -Name Path).Path -Split ";" -Replace "\\+$", ""'
    echo '$PathAlreadyExists = 0; foreach ($Path in $PathArray) { if ($Path -contains $PathToAdd ) { $PathAlreadyExists = 1 } }   # Use "-contains" for arrays'
    echo 'if ($PathAlreadyExists -eq 0) { $PathNew = $PathOld + ";" + $PathToAdd ; Set-ItemProperty -Path $RegistrySystemPath -Name Path -Value $PathNew }'
    ""
    $Paths = $env:Path.Split(';') | select -Unique | sort   # Array of current Paths
    # $CleanedInputList = @()
    # $NewPath | % { if (Test-Path $_) { 
    if ($NewPath -ne $null) {
        if (Test-Path $NewPath) { 
            $NewPaths = $Paths + $NewPath | select -Unique
            [Environment]::SetEnvironmentVariable( "Path", $($NewPaths -join ';'), [System.EnvironmentVariableTarget]::Machine )
            Write-Verbose "Processed the following new path (removed duplicates and bad paths):"
            Write-Verbose ($NEwPath | Out-String)
        }
    }
}

function AddTo-Path {
    param ( 
        [string]$PathToAdd,
        [Parameter(Mandatory=$true)][ValidateSet('System','User')]      [string]$UserType,
        [Parameter(Mandatory=$true)][ValidateSet('Path','PSModulePath')][string]$PathType
    )
    # https://stackoverflow.com/questions/714877/setting-windows-powershell-environment-variables/60636740#60636740

    # AddTo-Path "C:\XXX" 'System' "PSModulePath"
    if ($UserType -eq "User"   ) { $RegPropertyLocation = 'HKCU:\Environment' } # also note: Registry::HKEY_LOCAL_MACHINE\ format
    if ($UserType -eq "System" ) { $RegPropertyLocation = 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' }
    "`nAdd '$PathToAdd' (if not already present) into the $UserType `$$PathType"
    "The '$UserType' environment variables are held in the registry at '$RegPropertyLocation'"
    $PathOld = (Get-ItemProperty -Path $RegPropertyLocation -Name $PathType).$PathType
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

function aliases ($search) {
    ""
    ":: All aliases matching '$search' sorted by alias name:`n"
    $byalias = (alias "*$search*" | select Name,Definition | sort Name -Unique | % { $_.Name + " (" + $_.Definition +"), " } | Out-String) -replace "`r`n", ""
    Write-Wrap $byalias
    ""
    ":: All aliases matching '$search' sorted by Cmdlet name:`n"
    $bycmdlet = (alias "*$search*" | select Name,Definition | sort Definition -Unique | % { $_.Definition + " (" + $_.Name +"), " } | Out-String) -replace "`r`n", ""
    Write-Wrap $bycmdlet
    ""
}
# https://github.com/NightWolf92/Powershell-Repository
# https://github.com/wdomon/Script-Sharing
# https://www.reddit.com/r/PowerShell/comments/ejf8z4/script_sharing_just_some_scripts_i_wrote_and_use/
# I have been back and forth with putting all functions inside a file or a module over the years and I decided to just keep them in self contained scripts. Seems to work in the environments I have worked in as most admins don't know how to handle modules and prefer a single file. I use this template as you can see in my scripts here. For any changes, you can use Npp and bulk find and replace.
# Edit: I wholeheartedly agree that modules are better though, this method just works better when you work with admins that aren't very familiar with powershell.

function PSModulePath ($search) {
    ""
    ":: Complete `$env:PSModulePath string:`n"
    $env:PSModulePath
    $env:PSModulePath -Split ";" -replace "\\+$", "" | sort
    ""
    ""
    $PathArray = $env:PSModulePath -Split ";" -replace "\\+$", "" | sort
    ""
    foreach ($Path in $PathArray) {
        if ($Path -ne '') {
            $Path = $Path.TrimEnd('\')
            echo ":: Modules under '$Path':`n"
            # if (Test-Path $Path) {
            $temp = (dir "$Path\*$search*" -Directory -EA Silent | Select -ExpandProperty Name) -join ", "   # mods w*  => mods w** which resolves fine
            if ($temp -eq "") { "No matches for $search exist in this path" }
            # } else { "This path is in `$env:PSModulePath but the folder does not exist"}
            Write-Wrap $temp
            ""
        }
    }
    ""
}
Set-Alias mods PSModulePath -Description "Show all Modules installed on this system in each of the Module paths (`$Env:PSModulePath)."
# foreach ($Path in $PathArray) { (dir $Path | Select -ExpandProperty Name) -join ", " }
# Find all mods, Get-Command -Module $module for all
# List the path to the module, then the commands in there in comma separated form, like paths
# display some kind of table showing all of this
# if give a $module, then get as much info as possible for that module! is it installed, where is it located,
# what files are in that folder, what versions are available? is there a manifest (.psf1), how many functions
# are available.

function mod ($Module, $searchfor, [switch]$ShowModulesHere, [switch]$Info) {
    # First check if there is a folder in any module path, no point if not
    # If duplicates, resolve (this could be complex, various StackOverflows on this)
    # Function will perform an import only if the module is not loaded
    # foreach PSModulePath folder, check (Test-Path (Join-Path $AdminModulePath $Name)) {
    # -ShowModulesHere : what does this do, I forget
    # -Info : this should one line per function in Module and a quick description of the function.
    #    Also possibly sortable in various ways maybe?
    if ($null -eq $Module) { "You must specify a Module to examine. Run 'mods' to see available Modules." ; break }
    ""
    if ([bool](Get-Module $Module -ListAvailable) -eq $true) {
        if ([bool](Get-Module $Module) -eq $true) { "## Module '$Module' (version $ModuleVer) is already imported, so all functions are available." }
        else { "## Module '$Module' was not imported, running import now ..." ; Import-Module $Module }   # Do this first as $ModulePath will fail if not imported
        $ModulePath = ((Get-Module $Module | select Path).Path).TrimEnd('\')
        $ModuleVer = (Get-Module $Module -ListAvailable | select Version).Version | sls "\d"
    }
    else { "## Could not find Module '$Module' in any available Module folders:`n   $env:PSModulePath`n" ; break }
    
    # ":: '$Module' is version $($ModuleVer)`n   $ModulePath"
    
    $ModuleRoot = Split-Path ((Get-Module $Module | select Path).Path).TrimEnd("\")
    $ModulesHere = (dir $Path -Directory | Select -ExpandProperty Name) -join ", "
    
    if ($Info) {
        ""
        foreach ($i in (Get-Command -Module $Module).Name) { 
            $out = $i   # Parse the info string from after the "{" 
            $type = "" ; try { $type = ((gcm $i -EA silent).CommandType); } catch { $searchforerr = 1 }
            $out += "   # $type"
            $syntax = Get-Command $i -Syntax
            $searchforinition = "" ; if ($type -eq "Alias") { $searchforinition = (get-alias $i).Definition }
            $syntax = $syntax -replace $searchforinition, ""
            if ($type -eq "Alias") { $out += " for '$searchforinition'" }
            $out
            if ($type -eq "Function") { $syntax = $syntax -replace $i, "" }
            if ($type -eq "Cmdlet") { $syntax = $syntax -replace $i, "" }
            if (!([string]::IsNullOrWhiteSpace($syntax))) { 
                $syntax -split '\r\n' | where {$_} | foreach { "Syntax =>   $_" | Write-Wrap }
            }
            ""
        }
        ""
    }
    else {
        ""
        if ($null -ne $searchfor) {
            "## Searching '$Module' only for commands with '$searchfor' in the name:"
            $out = ""; foreach ($i in (Get-Command -Module $Module | ? Name -match $searchfor).Name) { $out += " $i," } ; "" ; Write-Wrap $out.TrimEnd(", ") ; ""
        }
        else {
            "## Module contents:"
            $out = ""; foreach ($i in (Get-Command -Module $Module).Name) { $out += " $i," } ; "" ; Write-Wrap $out.TrimEnd(", ") ; ""
        }
    }
    $ModPaths = $env:PSModulePath -Split ";" -replace "\\+$", "" | sort
    "## Module Paths (`$env:PSModulePath):"
    "   $env:PSModulePath"
    # foreach ($i in $ModPaths) { "   $i"}
    # ""
    # ":: '$Module' Path:`n`n   $ModulePath"
    # ""
    foreach ($i in $ModPaths) {
        if (!([string]::IsNullOrWhiteSpace($i))) {
           if ($ModulePath | sls $i -SimpleMatch) { $ModRoot = $i ; "`n## `$env:PSModulePath parent location is:`n   $i" }
        }
    }
    ""

    # if ($searchfor -ne $null) {
    #     ":: Press any key to open '$searchfor' definition:"
    #     pause
    #     ""
    #     def $searchfor
    #     ""
    # }

    if ($ShowModulesHere -eq $true) {
        "## This `$env:PSModulePath root also contains the following Modules:"
        ""
        (dir $ModRoot -Directory | Select -ExpandProperty Name) -join ", "
        ""
    }
}
function modi ($command) { mod $command -info | more }   # mod function with -i (info) switch

# Look for duplicates in modules folders
function Find-ModuleDuplicates {   # Broken WIP I think
    $hits = ""
    $ModPaths = $env:PSModulePath -Split ";" -replace "\\+$", "" | sort
    foreach ($i in $ModPaths) {
        foreach ($j in $ModPaths) {
            if ($j -notlike "*$i*") {
                $arr_i = (gci $i -Dir).Name
                $arr_j = (gci $j -Dir).Name
                foreach ($x in $arr_j) {
                    if ($arr_i -contains $x) {
                        $hits += "Module '$x' in '$i' has a duplicate`n" 
                    }
                }
            }
        }
    }
    if ($hits -ne "") { echo "" ; echo $hits }
    else { "`nNo duplicate Module folders were found`n" }
}




function SendEmail {
    param($server, $cpu, $mem, $disk) 
    # Code goes here to send email
    # https://stackoverflow.com/questions/60049690/need-to-monitor-server-resources-utilization-for-30-minutes-and-send-mail-in-pow/60050380#60050380
    # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.diagnostics/get-counter?view=powershell-7

    $cpuThreshold = 99
    $memThreshold = 99
    $diskThreshold = 100
    
    $jobHash = @{}
    $serverHash = @{}
    $serverList = @("localhost")
    
    foreach($server in $serverList) {
        $cpu = Start-Job -ComputerName $server -ScriptBlock { Get-Counter "\processor(_Total)\% Processor Time" -Continuous }
        $mem = Start-Job -ComputerName $server -ScriptBlock { Get-Counter -Counter "\Processor(_Total)\% Processor Time" -Continuous }
        $disk = Start-Job -ComputerName $server -ScriptBlock { Get-Counter -Counter "\LogicalDisk(C:)\% Free Space" -Continuous }
    
        $serverHash.Add("cpu", $cpu)
        $serverHash.Add("mem", $mem)
        $serverHash.Add("disk", $disk)
    
        $jobHash.Add($server, $serverHash)
    }
    
    Start-Sleep 10
    $totalLoops = 0
    
    while ($totalLoops -le 360) {
    
        foreach($server in $jobHash.Keys) {
            $cpu = (Receive-Job $jobHash[$server].cpu | % { (($_.readings.split(':'))[1]).Replace("`n","") } | measure -Maximum).Maximum
            $mem = (Receive-Job $jobHash[$server].mem | % { (($_.readings.split(':'))[1]).Replace("`n","") } | measure -Maximum).Maximum
            $disk = (Receive-Job $jobHash[$server].disk | % { (($_.readings.split(':'))[2]).Replace("`n","") } | measure -Maximum).Maximum
    
            if ($cpu -gt $cpuThreshold -or $mem -gt $memThreshold -or $disk -gt $diskThreshold) {
                Send-Email $server $cpu $mem $disk
            }
            Write-Output "CPU: $($cpu), Mem: $($mem), disk util: $($disk)"
        }
        Start-Sleep 1
        $totalLoops ++
    }
    
    Get-Job | Remove-Job
}


# ToDo: similar to sys (get-user, maybe rename sys to get-sys)
# Get-UserSession: returns currently running user sessions on a local or remote machine. Returns details about the session such as whether it is a console session, RDP, active, disconnected, or locked as well as the session duration, logon time, idle time (time since a keyboard or mouse was used), disconnect time and lock duration.
# Get-UserLogons: return the dates, times, and login method of user sessions for a local or remote machine.
# A general note on multi-threading for an example of CPU process collection:
# 1. This is slowers as it collects *all* processes first:
#    Get-Process | Where-Object {$_.ProcessName -eq 'svchost'} | Select-Object CPU | Where-Object {$_.CPU -ne $null}
# 2. This is faster as it filters only the processes you want to see, the does the manipulation then selects only what you need.
#    Get-Process -Name 'svchost' | Where-Object {$_.CPU -ne $null} | Select-Object CPU
# Get-Help *-Jobs to see all PSJobs Cmdlets
# All PSJobs are in one of eleven states. Most common are: Completed, Running, Blocked, Failed
# Get-Job to return status
# Start-Job -Scriptblock { Start-Sleep 5 }
# RunSpaces are better / faster than PowerShell Jobs, but are .NET so a little more complex.
# ForEach-Object -Parallel use runspaces to allow parallel foreach loops, but this is only available in PowerShell Core 7.
# https://adamtheautomator.com/powershell-multithreading/
# Split-Pipeline (PS v2.0 compatible) does similar to ForEach-Object -Parallel: https://github.com/nightroman/SplitPipeline
# PoshRSJob (use RunSpaces with same syntax as PS Jobs): https://github.com/proxb/PoshRSJob
# https://stackoverflow.com/questions/23327822/powershell-write-to-same-file-multiple-jobs
# This has some really important solutions:
# https://stackoverflow.com/questions/11973775/powershell-get-output-from-receive-job

function sys {
    $System = get-wmiobject -class "Win32_ComputerSystem"
    $Mem = [math]::Ceiling($System.TotalPhysicalMemory / 1024 / 1024 / 1024)
    
    # $wmi = gwmi -class Win32_OperatingSystem -computer "."   # Removed this method as not CIM compliant
    # $LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
    # [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
    $BootUpTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $CurrentDate = Get-Date
    $Uptime = $CurrentDate - $BootUpTime
    $s = "" ; if ($Uptime.Days -ne 1) {$s = "s"}
    $uptime_string = "$($uptime.days) day$s $($uptime.hours) hr $($uptime.minutes) min $($uptime.seconds) sec"

    # $temp_cpu = "$($env:TEMP)\ps_temp_cpu.txt"
    # $temp_cpu_cores = "$($env:TEMP)\ps_temp_cpu_cores.txt"
    # $temp_cpu_logical = "$($env:TEMP)\ps_temp_cpu_logical.txt"
    # rm -force $temp_cpu -EA silent ; rm -force $temp_cpu_cores -EA silent ; rm -force $temp_cpu_logical -EA silent
    # Start-Process -NoNewWindow -FilePath "powershell.exe" -ArgumentList "-NoLogo -NoProfile (Get-WmiObject -Class Win32_Processor).Name > $temp_cpu"
    # Start-Process -NoNewWindow -FilePath "powershell.exe" -ArgumentList "-NoLogo -NoProfile (Get-WmiObject -Class Win32_Processor).NumberOfCores > $temp_cpu_cores"
    # Start-Process -NoNewWindow -FilePath "powershell.exe" -ArgumentList "-NoLogo -NoProfile (Get-WmiObject -Class Win32_Processor).NumberOfLogicalProcessors > $temp_cpu_logical"
    # $job_cpu         = Start-Job -ScriptBlock { (Get-WmiObject -Class Win32_Processor).Name > $using:temp_cpu }
    # $job_cpu_cores   = Start-Job -ScriptBlock { (Get-WmiObject -Class Win32_Processor).NumberOfCores > $using:temp_cpu_cores }
    # $job_cpu_logical = Start-Job -ScriptBlock { (Get-WmiObject -Class Win32_Processor).NumberOfLogicalProcessors > $using:temp_cpu_logical }
    # If forget to use 'using' then everything looks like it is running, but will all fail.
    # Alternatively, can remove the output with $using and instead just run the job and receive the output and then do not need the temp files!
    # Receive-Job -Job $job_cpu -OutVariable job_cpu_output
    # https://stackoverflow.com/questions/25981724/automatically-removing-a-powershell-job-when-it-has-finished-asynchronously
    # https://powershell.org/forums/topic/self-terminating-jobs-in-powershell/
    $job_cpu         = Start-Job -ScriptBlock { (Get-WmiObject -Class Win32_Processor).Name }
    $job_cpu_cores   = Start-Job -ScriptBlock { (Get-WmiObject -Class Win32_Processor).NumberOfCores }
    $job_cpu_logical = Start-Job -ScriptBlock { (Get-WmiObject -Class Win32_Processor).NumberOfLogicalProcessors }
    ""
    "Hostname:        $($System.Name)"
    "Domain:          $($System.Domain)"
    "PrimaryOwner:    $($System.PrimaryOwnerName)"
    "Make/Model:      $($System.Manufacturer) ($($System.Model))"  #     "ComputerModel:  $((Get-WmiObject -Class:Win32_ComputerSystem).Model)"
    "SerialNumber:    $((Get-WmiObject -Class:Win32_BIOS).SerialNumber)"
    "PowerShell:      $($PSVersionTable.PSVersion)"
    "Windows Version: $($PSVersionTable.BuildVersion),   Windows ReleaseId: $((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'ReleaseId').ReleaseId)"
    "Display Card:    $((Get-WmiObject -Class:Win32_VideoController).Name)"
    "Display Driver:  $((Get-WmiObject -Class:Win32_VideoController).DriverVersion),   Description: $((Get-WmiObject -Class:Win32_VideoController).VideoModeDescription)"
    "Last Boot Time:  $([Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem | select 'LastBootUpTime').LastBootUpTime)),   Uptime: $uptime_string"

    # Note: -EA silent on Get-Item or will get an error
    Wait-Job $job_cpu         | Out-Null ; $job_cpu_out = Receive-Job -Job $job_cpu
    Wait-Job $job_cpu_cores   | Out-Null ; $job_cpu_cores_out = Receive-Job -Job $job_cpu_cores 
    Wait-Job $job_cpu_logical | Out-Null ; $job_cpu_logical_out = Receive-Job -Job $job_cpu_logical
    "CPU:             $job_cpu_out"
    "CPU Cores:       $job_cpu_cores_out,      CPU Logical Cores:   $job_cpu_logical_out"

    # Get-WmiObject -Class Win32_OperatingSystem | select @{N="LastBootTime"; E={$_.ConvertToDateTime($_.LastBootUpTime)}}
    # https://devblogs.microsoft.com/scripting/should-i-use-cim-or-wmi-with-windows-powershell/
    # $(wmic OS get LastBootupTime)

    # ipconfig | sls IPv4
    Get-Netipaddress | where AddressFamily -eq IPv4 | select IPAddress,InterfaceIndex,InterfaceAlias | sort InterfaceIndex
    $IPDefaultAddress = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].IPAddress[0]
    $IPDefaultGateway = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].DefaultIPGateway[0]
    "[Default IPAddress : $IPDefaultAddress / $IPDefaultGateway]" 
    ""
    
    # while (!(Test-Path $temp_cpu)) { while ((Get-Item $temp_cpu -EA silent).length -eq 0kb) { Start-Sleep -Milliseconds 500 } }
    # while (!(Test-Path $temp_cpu_cores)) { while ((Get-Item $temp_cpu_cores -EA silent).length -eq 0kb) { Start-Sleep -Milliseconds 500 } }
    # while (!(Test-Path $temp_cpu_logical)) { while ((Get-Item $temp_cpu_logical -EA silent).length -eq 0kb) { Start-Sleep -Milliseconds 500 } }
    # "CPU:       $(cat $temp_cpu_out)"
    # "CPU Cores: $(cat $temp_cpu_cores_out),   CPU Logical Cores: $(cat $temp_cpu_logical_out)"
    
    # while (!(Test-Path $temp_cpu)) { while ((Get-Item $temp_cpu -EA silent).length -eq 0kb) { Start-Sleep -Milliseconds 500 } }
    # "CPU:               $(cat $temp_cpu)"
    # while (!(Test-Path $temp_cpu_cores)) { while ((Get-Item $temp_cpu_cores -EA silent).length -eq 0kb) { Start-Sleep -Milliseconds 500 } }
    # "CPU Cores:         $(cat $temp_cpu_cores)"
    # while (!(Test-Path $temp_cpu_logical)) { while ((Get-Item $temp_cpu_logical -EA silent).length -eq 0kb) { Start-Sleep -Milliseconds 500 } }
    # "CPU Logical:       $(cat $temp_cpu_logical)"
    # Get-CimInstance -ComputerName localhost -Class CIM_Processor -ErrorAction Stop | Select-Object *
    # https://improvescripting.com/get-processor-cpu-information-using-powershell-script/

    # Get-PSDrive | sort -Descending Free | Format-Table
    # https://stackoverflow.com/questions/37154375/display-disk-size-and-freespace-in-gb
    # https://www.petri.com/checking-system-drive-free-space-with-wmi-and-powershell
    # https://www.oxfordsbsguy.com/2017/02/08/powershell-how-to-check-for-drives-with-less-than-10gb-of-free-diskspace/
    (gwmi win32_logicaldisk | Format-Table DeviceId, VolumeName, @{ n = "Size(GB)"; e = { [math]::Round($_.Size/1GB,2) } }, @{ n = "Free(GB)"; e = {[math]::Round($_.FreeSpace/1GB,2) } } | Out-String) -replace '(?m)^\r?\n'
    # gwmi win32_logicaldisk | Format-Table DeviceId, VolumeName, @{n="Size(GB)";e={[math]::Round($_.Size/1GB,2)}},@{n="Free(GB)";e={[math]::Round($_.FreeSpace/1GB,2)}}
    # Get-Volume | Where-Object {($_.SizeRemaining -lt 10000000000) -and ($_.DriveType -eq "FIXED") -and ($_.FileSystemLabel -ne "System Reserved")}
    
    # https://www.hanselman.com/blog/CalculateYourWEIWindowsExperienceIndexUnderWindows81.aspx
    # gwmi win32_winsat | select-object CPUScore,D3DScore,DiskScore,GraphicsScore,MemoryScore,TimeTaken,WinSATAssessmentState,WinSPRLevel,PSComputerName
    ((gwmi win32_winsat | select-object CPUScore,D3DScore,DiskScore,GraphicsScore,MemoryScore,WinSPRLevel | ft) | Out-String) -replace '(?m)^\r?\n'   # removed ,WinSATAssessmentState
    # $out = ""; foreach ($i in $WinSat_output) {$out += " $i,"} ; "" ; Write-Wrap $out.TrimEnd(", ")

    
    # https://superuser.com/questions/769679/powershell-get-list-of-folders-shared
    (get-WmiObject -class Win32_Share | ft | Out-String) -replace '(?m)^\r?\n'
}

function Get-PCInfo {
<# 
 .SYNOPSIS
  Function to ping and report on given one or more Windows computers.

 .DESCRIPTION
  Function to ping and report on given one or more Windows computers.
  If the computer has more than one network interface, this function will report all IP and MAC addresses

 .PARAMETER ComputerName
  One or more computer names to be reported on. This defaults to the current computer.

 .PARAMETER Cred
  PS Credential object that can be obtained from Get-Credential or Get-SBCredential

 .PARAMETER Refresh
  This switch will supress progress messages to speed up processing.

 .OUTPUTS 
  The function returns a PS object that has the following properties/example:
    ComputerName   : WIN10G2-Sam1
    Status         : Online
    IPAddress      : 192.168.214.118
    MACAddress     : 00:xx:xx:xx:xx:xx
    DateBuilt      : 9/6/2019 10:38:13 AM
    OSVersion      : 10.0.18363
    OSCaption      : Microsoft Windows 10 Enterprise
    OSArchitecture : 64-bit
    Model          : Virtual Machine
    Manufacturer   : Microsoft Corporation
    VM             : True
    LastBootTime   : 3/26/2020 9:38:45 PM

 .EXAMPLE
  Get-PCInfo 
  This returns the current PC information

 .EXAMPLE
  $PCInfo = Get-PCInfo -ComputerName @('PC1','PC2','PC3')
  This checks the listed computers and saves the collected information in $PCInfo variable

 .EXAMPLE
  (Import-Csv .\ComputerList1.csv).ComputerName | Get-PCInfo | Export-Csv .\ComputerReport.csv -NoType
  This example will read a list of computer names from the CSV file provided which has a 'ComputerName' column,
  gather each computer information and save it to the provided CSV output file.

 .EXAMPLE
  Get-PCInfo -ComputerName Server111 -Cred (Get-SBCredential 'domain\user')
  This example will report on information of the provided computer using the provided credentials

 .LINK 
  https://superwidgets.wordpress.com/2017/01/04/powershell-script-to-report-on-computer-inventory/

 .NOTES
  Function by Sam Boutros
    31 October 2014 v0.1 
    4  January 2017 v0.2
    17 March   2017 v0.3 - chnaged the logic to output 1 record per computer even when it has several NICs
    2  April   2020 v0.4 - Added Silent switch to speed up processing of large number of computers
        Switched to using Get-SBWMI instead of Get-WMIObject
        Added Cred Parameter to be able to query computers outside the domain
    16 December 2021 v0.5 - Added FreeRAM (percent) and CPU (used percent) metrics.

#>

    [CmdletBinding(ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
            [String[]]$ComputerName = $env:COMPUTERNAME,
        [Parameter(Mandatory=$false)][PSCredential]$Cred,
        [Parameter(Mandatory=$false)][Switch]$Silent
    )

    Begin { }

    Process {
        
        foreach ($PC in $ComputerName) {
            if (-not $Silent) { Write-Log 'Checking computer',$PC Green,Cyan -NoNewLine }
            
            try {
                $Result = Test-Connection -ComputerName $PC -Count 2 -ErrorAction Stop 
                if ($Cred) {
                    $OS  = Get-SBWMI -ComputerName $PC -Class Win32_OperatingSystem -Cred $Cred -EA 0
                    $Mfg = Get-SBWMI -ComputerName $PC -Class Win32_ComputerSystem  -Cred $Cred -EA 0
                    $CPU = Get-SBWMI -ComputerName $PC -Class Win32_Processor       -Cred $Cred -EA 0
                    $IPs = (Get-SBWMI -ComputerName $PC -Class Win32_NetworkAdapterConfiguration -Cred $Cred -EA 0 | 
                            Where { $_.IpEnabled }).IPAddress | where { $_ -match "\." } # IPv4 only
                } else {
                    $OS  = Get-SBWMI -ComputerName $PC -Class Win32_OperatingSystem -EA 0
                    $Mfg = Get-SBWMI -ComputerName $PC -Class Win32_ComputerSystem  -EA 0
                    $CPU = Get-SBWMI -ComputerName $PC -Class Win32_Processor       -EA 0
                    $IPs = (Get-SBWMI -ComputerName $PC -Class Win32_NetworkAdapterConfiguration -EA 0 | 
                            Where { $_.IpEnabled }).IPAddress | where { $_ -match "\." } # IPv4 only
                }
                $MACs = foreach ($IPAddress in $IPs) {
                    if ($Cred) {
                        (Get-SBWMI -ComputerName $PC -Class Win32_NetworkAdapterConfiguration -Cred $Cred -EA 0 | 
                            Where { $_.IPAddress -eq $IPAddress }).MACAddress
                    } else {
                        (Get-SBWMI -ComputerName $PC -Class Win32_NetworkAdapterConfiguration -EA 0 | 
                            Where { $_.IPAddress -eq $IPAddress }).MACAddress
                    }                        
                }
                if (-not $Silent) { Write-Log 'done' Green }
                [PSCustomObject]@{
                    ComputerName   = $PC
                    Status         = 'Online'
                    IPAddress      = $IPs -join ', '
                    MACAddress     = $MACs -join ', '
                    DateBuilt      = ([WMI]'').ConvertToDateTime($OS.InstallDate)
                    OSVersion      = $OS.Version
                    OSCaption      = $OS.Caption
                    OSArchitecture = $OS.OSArchitecture
                    Model          = $Mfg.model
                    Manufacturer   = $Mfg.Manufacturer
                    VM             = $(if ($Mfg.Manufacturer -match 'vmware' -or $Mfg.Manufacturer -match 'microsoft') { $true } else { $false })
                    LastBootTime   = ([WMI]'').ConvertToDateTime($OS.LastBootUpTime)
                    FreeRAM        = 100 - [math]::Round(($OS.FreePhysicalMemory/$OS.TotalVisibleMemorySize)*100,0)
                    CPU            = [math]::Round(($CPU | measure LoadPercentage -Average).Average,0)
                }
            } catch { # either ping failed or access denied 
                if ($Result) {
                    if (-not $Silent) { Write-Log 'done' Magenta }
                    [PSCustomObject]@{
                        ComputerName   = $PC
                        Status         = $Error[0].Exception
                    }
                } else {
                    if (-not $Silent) { Write-Log 'done' Yellow }
                    [PSCustomObject]@{
                        ComputerName   = $PC
                        Status         = 'No response to ping'
                    }
                }
            }
        }
    }

    End { }
}


##### Registry #####
# https://gallery.technet.microsoft.com/scriptcenter/Get-RegistryKeyLastWriteTim-63f4dd96

function Add-RegistryFavorites {
    # Following is a .reg from the current user Favorites. Should populate this with useful locations in the below format
    # Windows Registry Editor Version 5.00
    # [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites]
    # "Environment (User)"="Computer\\HKEY_CURRENT_USER\\Environment"

    # Not registry, but keep this here to rebuild if required:
    ##### C:\Users\Public\Desktop\desktop.ini
    # [.ShellClassInfo]
    # LocalizedResourceName=@%SystemRoot%\system32\shell32.dll,-21799
    ##### C:\Users\Boss\Desktop\desktop.ini
    # [.ShellClassInfo]
    # LocalizedResourceName=@%SystemRoot%\system32\shell32.dll,-21769
    # IconResource=%SystemRoot%\system32\imageres.dll,-183
}

################
# Benchmarking #
################
#
# https://www.powershelladmin.com/wiki/PowerShell_benchmarking_module_built_around_Measure-Command
# https://improvescripting.com/how-to-benchmark-scripts-with-powershell/
# Install-Module -Name Benchmark -RequiredVersion 1.2.2
# measure / difference datetime / stopwatch:  https://www.pluralsight.com/blog/tutorials/measure-powershell-scripts-speed


# https://stackoverflow.com/questions/49011865/powershell-script-to-check-disk-free-space
# $DiskReport | Select-Object @{Label = "Server Name";Expression = {$_.SystemName}},
#     @{Label = "Drive Letter";Expression = {$_.DeviceID}},
#     @{Label = "Total Capacity (GB)";Expression = {"{0:N1}" -f( $_.Size / 1gb)}},
#     @{Label = "Free Space (GB)";Expression = {"{0:N1}" -f( $_.Freespace / 1gb ) }},
#     @{Label = 'Free Space (%)'; Expression = {"{0:P0}" -f ($_.freespace/$_.size)}} |
#     Export-Csv -path "c:\data\server\ServerStorageReport\DiskReport\DiskReport_$logDate.csv" -NoTypeInformation


# https://stackoverflow.com/questions/185575/powershell-equivalent-of-bash-ampersand-for-forking-running-background-proce
# Start-Process -NoNewWindow ping google.com
# You can also add this as a function in your profile:   function bg() {Start-Process -NoNewWindow @args}
# and then the invocation becomes:   bg ping google.com
# In my opinion, Start-Job is an overkill for the simple use case of running a process in the background:
# 
#     Start-Job does not preserve the current directory (because it runs in a separate session). You cannot do "Start-Job {notepad myfile.txt}" where myfile.txt is in the current directory.
#     The output is not displayed automatically. You need to run Receive-Job with the ID of the job as parameter.
# 
# NOTE: Regarding your initial example, "bg sleep 30" would not work because sleep is a Powershell commandlet. Start-Process only works when you actually fork a process.


# http://powershell-guru.com/powershell-tip-7-convert-wmi-date-to-datetime/
# https://stackoverflow.com/questions/17681234/how-do-i-get-total-physical-memory-size-using-powershell-without-wmi
# https://devblogs.microsoft.com/scripting/use-powershell-and-wmi-to-get-processor-information/
# $WinVer = New-Object -TypeName PSObject   # Using a PSObject construct
# $WinVer | Add-Member -MemberType NoteProperty -Name Major -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' CurrentMajorVersionNumber).CurrentMajorVersionNumber
# $WinVer | Add-Member -MemberType NoteProperty -Name Minor -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' CurrentMinorVersionNumber).CurrentMinorVersionNumber
# $WinVer | Add-Member -MemberType NoteProperty -Name Build -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' CurrentBuild).CurrentBuild
# $WinVer | Add-Member -MemberType NoteProperty -Name Revision -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' UBR).UBR
# $WinVer
#
# More advanced techniques here: https://lazywinadmin.com/2015/03/standard-and-advanced-powershell.html
# Get-WMIObject -Class Win32_ComputerSystem # Information about the System
# Get-WMIObject -Class Win32_BIOS           # Information about the BIOS
# Get-WMIObject -Class Win32_Baseboard      # Information about the Motherboard
# Get-WMIObject -Class Win32_Processor      # Information about the CPU
# Get-WMIObject -Class Win32_LogicalDisk    # Information about Logical Drives (Includes mapped drives and I believe PSDrives)
# Get-WMIObject -Class Win32_DiskDrive      # Information about Physical Drives
# Get-WMIObject -Class Win32_PhysicalMemory # Information about the Memory
# Get-WMIObject -Class Win32_NetworkAdapter # Information about the NIC
# Get-WMIObject -Class Win32_NetworkAdapterConfiguration # Information about the NIC's Configuration
# You can also shortform the command from;
# Get-WMIObject -Class Win32_ComputerSystem into gwmi Win32_ComputerSystem
# Get-NetAdapter | Restart-NetAdapter   # Restart all network adapters

# $names = Get-Content "C:\scripts\servers.txt"
# @(
#  foreach ($name in $names) {
#      if (Test-Connection -ComputerName $name -Count 1 -ErrorAction SilentlyContinue) {
#          $wmi = gwmi -class Win32_OperatingSystem -computer $name
#          $LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
#          [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
#          Write-output "$name Uptime is  $($uptime.days) Days $($uptime.hours) Hours $($uptime.minutes) Minutes $($uptime.seconds) Seconds"
#      } else {
#          Write-output "$name is not pinging"
#      }
#  }
# ) | Out-file -FilePath "c:\chethan\results.txt"



function m ($cmd, [switch]$Definition, [switch]$Examples, [switch]$Synopsis, [switch]$Syntax) {
    # m (man), explore and open about_ Topics
    # Define this inside the function so that the main function is self-contained
    function Write-Wrap {
        [CmdletBinding()]Param( [parameter(Mandatory=1, ValueFromPipeline=1, ValueFromPipelineByPropertyName=1)] [Object[]]$chunk )
        PROCESS {
            $Lines = @()
            foreach ($line in $chunk) {
                $str = ''; $counter = 0
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

    if ($definition -eq $true) { def $cmd ; break }
    # .NET operations
    if ($cmd -eq '.net' -or $cmd -eq 'dotnet' -or $cmd -eq 'system.string' -or $cmd -eq '.') {
        Write-Host "`n.NET methods and properties available (i.e. from .NET, not part of PowerShell)"
        Write-Host "https://docs.microsoft.com/en-GB/dotnet/api/System.String?view=netframework-4.8"
        Write-Host "Expand the Properties and Methods section in this link for more information."
        Write-Host "`nFor strings:"
        Write-Wrap ("" | get-member | select Name | % { "." + $_.Name + "," } | Out-String).replace("about_", "").replace("`r`n", " ").trim(", ")
        Write-Host "`nFor integers:"
        Write-Wrap (0 | get-member | select Name | % { "." + $_.Name + "," } | Out-String).replace("about_", "").replace("`r`n", " ").trim(", ")
        Write-Host ""
        break
    }
    # if ($cmd -eq '.Split') { "".split }   # > "$($env:TEMP)\def.txt"
    # Should be able to do -match "\.*" then use that to build the "".split
    # "".split
    # OverloadDefinitions
    # -------------------
    # [Enum]::GetNames([StringSplitOptions])
    # None
    # RemoveEmptyEntries
    # https://stackoverflow.com/questions/59212258/get-help-for-trim-trim-replace-replace-split-split-and-other-strin/59324892#59324892

    # Some way to remove all of the below and replace with iterating through
    # if ($cmd -eq '.net' -or $cmd -eq 'dotnet' -or $cmd -eq 'system.string') {
    #     Write-Host "`nSystem.String (with a string):"
    #     Write-Wrap ("" | get-member | select Name | % { "." + $_.Name + "," } | Out-String).replace("about_", "").replace("`r`n", " ").trim(", ")
    
    # https://docs.microsoft.com/en-GB/dotnet/api/System.String.Clone?view=netframework-4.8
    # https://docs.microsoft.com/en-GB/dotnet/api/System.String.<XXX>?view=netframework-4.8
    # https://docs.microsoft.com/en-GB/dotnet/api/System.String?view=netframework-4.8
    # Show-TypeHelp.ps1 string contains
    # Cannot add 'break' at end of these as they return nothing and will exit the console if so

    if ($cmd -ne $null) { if ($cmd.StartsWith(".")) {
        if ($cmd -eq '.Clone') { "".clone ; "Returns a reference to this instance of String." ; return }
        if ($cmd -eq '.CompareTo') { "".compareto
Write-Wrap @"
Compares this instance with a specified object or String and returns an integer that indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified object or String.
    
CompareTo(String)	
Compares this instance with a specified String object and indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified string.
    
"@ ; return }
        if ($cmd -eq '.Contains') { "".contains
@"
Returns a value indicating whether a specified substring occurs within this string.

Contains(String)
Returns a value indicating whether a specified substring occurs within this string.
"@ ; return }
        if ($cmd -eq '.CopyTo') { "".copyto ; return }
        if ($cmd -eq '.EndsWith') { "".endswith
@"
Determines whether the end of this string instance matches a specified string.

EndsWith(String)
Determines whether the end of this string instance matches the specified string.

EndsWith(String, StringComparison)	
Determines whether the end of this string instance matches the specified string when compared using the specified comparison option.

EndsWith(String, Boolean, CultureInfo)	
Determines whether the end of this string instance matches the specified string when compared using the specified culture.
"@ ; return }
        if ($cmd -eq '.Equals') { "".equals ; return }
        if ($cmd -eq '.GetEnumerator') { "".getenumerator ; return }
        if ($cmd -eq '.GetHashCode') { "".gethashcode ; return }
        if ($cmd -eq '.GetType') { "".gettype ; return }
        if ($cmd -eq '.GetTypeCode') { "".gettypecode ; return }
        if ($cmd -eq '.IndexOf') { "".indexof ; return }
        if ($cmd -eq '.IndexOfAny') { "".indexofany ; return }
        if ($cmd -eq '.Insert') { "".insert ; return }
        if ($cmd -eq '.IsNormalized') { "".isnormalized ; return }
        if ($cmd -eq '.LastIndexOf') { "".lastindexof ; return }
        if ($cmd -eq '.LastIndexOfAny') { "".lastindexofany ; return }
        if ($cmd -eq '.Normalize') { "".normalize ; return }
        if ($cmd -eq '.PadLeft') { "".padleft ; return }
        if ($cmd -eq '.PadRight') { "".padright ; return }
        if ($cmd -eq '.Remove') { "".remove ; return }
        if ($cmd -eq '.Replace') { "".replace ; return }
        if ($cmd -eq ".Split") { "".split
@"
StringSplitOptions, RemoveEmptyEntries omit empty array elements from the array returned; or None to include empty array elements in the array returned.
[Enum]::GetNames([StringSplitOptions]) 
"@ ; return }
        if ($cmd -eq '.StartsWith') { "".startswith ; return }
        if ($cmd -eq '.Substring') { "".substring ; return }
        if ($cmd -eq '.ToBoolean') { "".toboolean ; return }
        if ($cmd -eq '.ToByte') { "".tobyte ; return }
        if ($cmd -eq '.ToChar') { "".tochar ; return }
        if ($cmd -eq '.ToCharArray') { "".tochararray ; return }
        if ($cmd -eq '.ToDateTime') { "".todatetime ; return }
        if ($cmd -eq '.ToDecimal') { "".todecimal ; return }
        if ($cmd -eq '.ToDouble') { "".todouble ; return }
        if ($cmd -eq '.ToInt16') { "".toint16 ; return }
        if ($cmd -eq '.ToInt32') { "".toint32 ; return }
        if ($cmd -eq '.ToInt64') { "".toint64 ; return }
        if ($cmd -eq '.ToLower') { "".tolower ; return }
        if ($cmd -eq '.ToLowerInvariant') { "".tolowerinvariant ; return }
        if ($cmd -eq '.ToSByte') { "".tosbyte ; return }
        if ($cmd -eq '.ToSingle') { "".tosingle ; return }
        if ($cmd -eq '.ToString') { "".tostring ; return }
        if ($cmd -eq '.ToType') { "".totype ; return }
        if ($cmd -eq '.ToUInt16') { "".toint16 ; return }
        if ($cmd -eq '.ToUInt32') { "".toint32 ; return }
        if ($cmd -eq '.ToUInt64') { "".touint64 ; return }
        if ($cmd -eq '.ToUpper') { "".toupper ; return }
        if ($cmd -eq '.ToUpperInvariant') { "".toupperinvariant ; return }
        if ($cmd -eq '.Trim') { "".trim ; return }
        if ($cmd -eq '.TrimEnd') { "".trimend ; return }   # ("xxxyyy", "yyyxxx").trimend("x")   will delete multiple characters if same, not use with strings or array
        if ($cmd -eq '.TrimStart') { "".trimstart ; return }
        if ($cmd -eq '.Chars') { "".chars ; return }
        if ($cmd -eq '.Length') { "".length ; return }
        return 
    } }

    # reserved keywords
    if ($cmd -eq 'begin' -or $cmd -eq 'break' -or $cmd -eq 'catch' -or $cmd -eq 'class' -or $cmd -eq 'continue' -or $cmd -eq 'data' -or
        $cmd -eq 'define' -or $cmd -eq 'do' -or $cmd -eq 'dynamicparam' -or $cmd -eq 'else' -or $cmd -eq 'elseif' -or $mcd -eq 'end' -or
        $cmd -eq 'exit' -or $cmd -eq 'filter' -or $cmd -eq 'finally' -or $cmd -eq 'for' -or $cmd -eq 'foreach' -or $cmd -eq 'from' -or
        $cmd -eq 'function' -or $cmd -eq 'if' -or $cmd -eq 'in' -or $cmd -eq 'inlinescript' -or $cmd -eq 'parallel' -or $cmd -eq 'param' -or
        $cmd -eq 'process' -or $cmd -eq 'return' -or $cmd -eq 'sequence' -or $cmd -eq 'switch' -or $cmd -eq 'throw' -or $cmd -eq 'trap' -or
        $cmd -eq 'try' -or $cmd -eq 'until' -or $cmd -eq 'using' -or $cmd -eq 'var' -or $cmd -eq 'while' -or $cmd -eq 'workflow' )
    {
        ""
        Write-Wrap "Note: The search term '$cmd' is a Reserved Word that has a special meaning in PowerShell."
        ""
        Write-Wrap "'m about_Reserved_Words' shows an overview of the reserved words."
        Write-Wrap "'m about_Language_Keywords' will list the topic that contains detailed information on a given keyword."
        ""
    }

    # Current set of reserved words that are in use:
    # Begin              Exit               Process
    # Break              Filter             Return
    # Catch              Finally            Sequence
    # Class              For                Switch
    # Continue           ForEach            Throw
    # Data               From               Trap
    # Define             Function           Try
    # Do                 If                 Until
    # DynamicParam       In                 Using
    # Else               InlineScript       Var
    # ElseIf             Parallel           While
    # End                Param              Workflow

    # Set of words reserved for future (but not currently in use):
    # assembly           module
    # base               namespace
    # command            private
    # configuration      static
    # enum               type
    # interface          

    if ($cmd -notmatch "^about_") {
        # { $cmd = $cmd -replace "about_", "" }  # Make input consistent with just the topic name
        # PowerShell Operators
        if ($cmd -eq 'ma' -or $cmd -eq 'about') { ma ; $foundmethod = $true }

        if ($cmd -eq '%') { "Used as both Arithmetic Operator and as an alias for Where-Object" ; pause }

        if ($cmd -eq 'split' -or $cmd -eq '-split') { Get-Help about_Split | more ; $foundmethod = $true ; return }
        if ($cmd -eq 'join' -or $cmd -eq '-join') { Get-Help about_Join | more ; $foundmethod = $true ; return }
        if ($cmd -eq 'redirect' -or $cmd -eq 'redirection') { Get-Help about_Redirection | more ; $foundmethod = $true ; return }
        if ($cmd -eq '>' -or $cmd -eq '>>' -or $cmd -eq '2>' -or $cmd -eq '2>&1') { Get-Help about_Redirection | more ; $foundmethod = $true ; return }
        if ($cmd -eq '@' -or $cmd -eq 'splat' -or $cmd -eq 'splatting') { Get-Help about_Splatting | more ; $foundmethod = $true ; return }
        if ($cmd -eq "-" -or $cmd -eq '*' -or $cmd -eq '/' -or $cmd -eq '+') { Get-Help about_Arithmetic_Operators | more ; $foundmethod = $true ; return }   # remove %, clash with ForEach-Object
        if ($cmd -eq "!" -or $cmd -eq 'and' -or $cmd -eq 'or' -or $cmd -eq 'xor' -or $cmd -eq 'not') { Get-Help about_Logical_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '-and' -or $cmd -eq '-or' -or $cmd -eq '-xor' -or $cmd -eq '-not') { Get-Help about_Logical_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq "+=" -or $cmd -eq '-=' -or $cmd -eq '*=' -or $cmd -eq '/=' -or $cmd -eq '%=' -or $cmd -eq '++' -or $cmd -eq '--') { Get-Help about_Assignment_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq "as" -or $cmd -eq '-as' -or $cmd -eq 'isnot' -or $cmd -eq '-isnot') { Get-Help about_Type_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '-eq' -or $cmd -eq '-ne' -or $cmd -eq '-gt' -or $cmd -eq '-ge' -or $cmd -eq '-lt' -or $cmd -eq '-le') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq 'eq' -or $cmd -eq 'ne' -or $cmd -eq 'gt' -or $cmd -eq 'ge' -or $cmd -eq 'lt' -or $cmd -eq 'le') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq 'like' -or $cmd -eq 'notlike' -or $cmd -eq 'match' -or $cmd -eq 'notmatch') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '-like' -or $cmd -eq '-notlike' -or $cmd -eq '-match' -or $cmd -eq '-notmatch') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq 'replace' -or $cmd -eq 'in' -or $cmd -eq 'notIn') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '-replace' -or $cmd -eq '-in' -or $cmd -eq '-notIn') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq 'contains' -or $cmd -eq 'notContains' -or $cmd -eq '-contains' -or $cmd -eq '-notContains') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '+' -or $cmd -eq '-' -or $cmd -eq '*' -or $cmd -eq '/' -or $cmd -eq '%') { Get-Help about_Arithmetic_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq 'shl' -or $cmd -eq 'shr' -or $cmd -eq '-shl' -or $cmd -eq '-shr') { Get-Help about_Arithmetic_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq "bAND" -or $cmd -eq 'bOR' -or  $cmd -eq 'bXOR' -or $cmd -eq 'bNOT') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq "-bAND" -or $cmd -eq '-bOR' -or  $cmd -eq '-bXOR' -or $cmd -eq '-bNOT') { Get-Help about_Comparison_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '@()' -or $cmd -eq '()' -or $cmd -eq '.' -or $cmd -eq '&') { Get-Help about_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '&&' -or $cmd -eq '||' -or $cmd -eq '|' -or $cmd -eq '::') { Get-Help about_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '[]' -or $cmd -eq ',' -or $cmd -eq '..' -or $cmd -eq '-f') { Get-Help about_Operators | more ; $foundmethod = $true ; return }
        if ($cmd -eq '?:' -or $cmd -eq '??' -or $cmd -eq '$()') { Get-Help about_Operators | more ; $foundmethod = $true ; return }
        # Reserved Keywords. The following reserved words don't appear anywhere else: Define, Using, Var
        if ($cmd -eq 'begin' -or $cmd -eq 'filter' -or $cmd -eq 'function' -or $cmd -eq 'param' -or $cmd -eq 'process' -or $cmd -eq 'end') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Functions, about_Functions_Advanced`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'try' -or $cmd -eq 'catch' -or $cmd -eq 'finally') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Try_Catch_Finally`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'break' -or $cmd -eq 'continue' -or $cmd -eq 'trap' -or $cmd -eq 'return') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Break, about_Continue, about_Trap, about_Try_Catch_Finally, about_Return`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'do' -or $cmd -eq 'until' -or $cmd -eq 'while') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Brexk, about_Do, about_While`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'data') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Data_Sections`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'DynamicParam') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Functions_Advanced_Parameters`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'if' -or $cmd -eq 'elseif' -or $cmd -eq 'else') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_If`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'exit') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nMain topic in here (about_Language_Keywords)`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'for' -or $cmd -eq 'foreach' -or $cmd -eq 'in') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_For, about_ForEach, about_In`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'from') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nReserved for future use`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'InlineScript') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_InlineScript`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'Class' -or $cmd -eq 'Hidden') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Classes, about_Hidden`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'Sequence') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Sequence`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'Switch') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Switch`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'Throw') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Throw, about_Functions_Advanced_Methods`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'Workflow') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Workflow`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'ForEach-Parallel' -or $cmd -eq 'Parallel') { Get-Help about_Language_Keywords | more ; Write-Host "`n`nAlso see: about_Parallel, about_ForEach-Parallel`n`n" -F Green ; $foundmethod = $true ; return }
        if ($cmd -eq 'Define' -or $cmd -eq 'Using' -or $cmd -eq 'Var') { Write-Host "`n`nThis reserved keyword is not currently used`n`n" -F Green ; $foundmethod = $true ; return }

        # m break fails to open about_break! why not
        # m splat pulls up about_Splatting twice in a row!
        # No about_ Topics found for 'define'. Searching among help / alias entries ...
        # Id     Name            PSJobTypeName   State         HasMoreData     Location             Command
        # --     ----            -------------   -----         -----------     --------             -------
        # 1      Job1            BackgroundJob   Running       True            localhost             if (-not (Test-Path $...
        # cat : Cannot find path 'C:\Users\Boss\AppData\Local\Temp\ps_help_synopsis.txt' because it does not exist.
        # At C:\Program Files\WindowsPowerShell\Modules\Custom-Tools\Custom-Tools.psm1:2392 char:25
        # +             $out_synopsis = cat $help_synopsis
        # +                         ~~~~~~~~~~~~~~~~~~
        #     + CategoryInfo          : ObjectNotFound: (C:\Users\Boss\A...lp_synopsis.txt:String) [Get-Content], ItemNotFoundException
        #     + FullyQualifiedErrorId : PathNotFound,Microsoft.PowerShell.Commands.GetContentCommand
        # 
        # Nothing was found from 'help *' ...
        # Last attempt is to directly try:
        #    help define
    }

    # Expand Synopsis: https://stackoverflow.com/questions/9775272/how-to-access-noteproperties-on-the-inputobject-to-a-remoting-session
    # For abount_ files, 'Synopsis' is a NoteProperty and only returns the first line (even with -ExpandProperty)
    # [array]$x = $profile | select -ExpandProperty 
    # [array]$x = get-help about_wql | select -ExpandProperty Synopsis
    # for ($i = 0; $i -lt $x.Count; $i ++)
    # {
    #     Write-Host -ForeGroundColor "Magenta" $i
    #     $x[$i]
    # }
    
    $about_name = "$($env:TEMP)\ps_about_name.txt"
    $about_namecsv = "$($env:TEMP)\ps_about_name.csv"
    $about_synopsis = "$($env:TEMP)\ps_about_synopsis.txt"
    if ((Get-Item $about_name -EA silent).length -eq 0kb) { rm -force $about_name -EA silent }
    if ((Get-Item $about_namecsv -EA silent).length -eq 0kb) { rm -force $about_namecsv -EA silent }
    if ((Get-Item $about_synopsis -EA silent).length -eq 0kb) { rm -force $about_synopsis -EA silent }
    if (-not (Test-Path $about_name)) { (get-help about_* | select Name | sort Name -Unique | Out-String) -replace "[Aa]bout_", "" > $about_name }   # $using: if in Job
    if (-not (Test-Path $about_namecsv)) { ((get-help about_* | select Name | sort Name -Unique | % { $_.Name + "," } | Out-String) -replace "[Aa]bout_", "" -replace "`r`n", " ").trim(", ") > $about_namecsv }
    if (-not (Test-Path $about_synopsis)) { (get-help about_* | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } | Out-String) -replace "[Aa]bout_", "" > $about_synopsis }
    $help_name = "$($env:TEMP)\ps_help_name.txt"
    $help_namecsv = "$($env:TEMP)\ps_help_name.csv"
    $help_synopsis = "$($env:TEMP)\ps_help_synopsis.txt"
    if ((Get-Item $help_name -EA silent).length -eq 0kb) { rm -force $help_name -EA silent }
    if ((Get-Item $help_namecsv -EA silent).length -eq 0kb) { rm -force $help_namecsv -EA silent }
    if ((Get-Item $help_synopsis -EA silent).length -eq 0kb) { rm -force $help_synopsis -EA silent }
    if (-not (Test-Path $help_name)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name | sort Name -Unique > $help_name }
    if (-not (Test-Path $help_namecsv)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name | sort Name -Unique | % { $_.Name + "," } > $help_namecsv }
    # $job_help_synopsis = Start-Job { if (-not (Test-Path $using:help_synopsis)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } > $using:help_synopsis } }

    # $about = get-help about_* | select Name,Synopsis
    # $job_about_name = Start-Job { if (-not (Test-Path $using:about_name)) { ($using:about | select Name | sort Name -Unique | Out-String) -replace "[Aa]bout_", "" > $using:about_name } }
    # $job_about_namecsv = Start-Job { if (-not (Test-Path $using:about_namecsv)) { (($using:about | select Name | sort Name -Unique | % { $_.Name + "," } | Out-String) -replace "[Aa]bout_", "" -replace "`r`n", " ").trim(", ") > $using:about_namecsv } }
    # $job_about_synopsis = Start-Job { if (-not (Test-Path $using:about_synopsis)) { ($using:about | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } | Out-String) -replace "[Aa]bout_", "" > $using:about_synopsis } }
    # $job_help_name = Start-Job { if (-not (Test-Path $using:help_name)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name | sort Name -Unique > $using:help_name } }
    # $job_namecsv !!!!!!!! = Start-Job { if (-not (Test-Path $using:help_name)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name | sort Name -Unique > $using:help_name } }
    # $job_help_synopsis = Start-Job { if (-not (Test-Path $using:help_synopsis)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } > $using:help_synopsis } }
    # Wait-Job $job_about_name, $job_about_namecsv, $job_about_synopsis, $job_help_name, $job_help_synopsis | Out-Null

    $out_name = cat $about_name
    $out_synopsis = cat $about_synopsis
    $out_csv = cat $about_namecsv

    if ($cmd -eq $null) {

        # $out_name = cat $about_name   # $topics = get-help about_* | select Name,Synopsis
        # $out_synopsis = cat $about_synopsis
        # $csv = cat $about_namecsv

        Write-Host ""
        Write-Host ""
        Write-Host ':::::::::::::::::::::::::::::::'
        Write-Host "::  about_ Topic & Synopsis  ::"
        Write-Host ':::::::::::::::::::::::::::::::'
        Write-Host ""
        Write-Wrap $out_synopsis
        Write-Host ""
        Write-Host ""
        Write-Host ':::::::::::::::::::::::::::::'
        Write-Host "::  about_ Topic csv List  ::"
        Write-Host ':::::::::::::::::::::::::::::'
        Write-Host "Total number of about_ Topics currently found: $($out_name.count)"
        Write-Host ""
        Write-Wrap $out_csv
        Write-Host ""
        Write-Host ""
        Write-Host ""
        Write-Host ':::::::::::::::::::::::::'
        Write-Host '::  m <Search-String>  ::'
        Write-Host ':::::::::::::::::::::::::'
        Write-Host ""
        Write-Host 'Expanded help tool with improved about_ Topic, .NET method, and PowerShell operator information more easily accessible.'   # mh should be "MAN / HELP"
        Write-Host ""
        $out_help = "The tool will cache some help details in the users temp folder (type 'm' on its own to update these, and 'dir `$env:temp\ps*' to view). "
        $out_help += "'m' uses the <Search-String> to find matching help information. This is done in a few steps to identify the most relevant information."
        Write-Wrap $out_help
        Write-Host ""
        Write-Wrap "1. First test if <Search-String> is a .NET string operator. e.g. m .clone or m .split. To list all string operators, use 'm .net'."
        Write-Host ""
        Write-Wrap "2. Look for PowerShell reserved keyboards / operators and open associated about_ topic file."
        Write-Wrap "e.g. 'm begin', 'm eq', 'm `"-and`"', 'm `"&`"', 'm Reserved_Words', 'm Language_Keywords', 'm Operators'"
        Write-Host ""
        Write-Wrap "3. Then, use <Search-String> to look for about_ Topics. To see all available about_ topics, just type 'm' (this will also offer to update the cached files and to run a full Update-Help if Admin)."
        Write-Host ""
        Write-Wrap "4. Then use <Search-String> to look for matching help information for a matching cmdlet or alias etc and will expand it using 'more' and -detail for help files so that the most relevant information is displayed."
        Write-Host ""
        Write-Wrap "If there are multiple topics matching <Search-String> the tool will display a list of those items. Refine <Search-String> and rerun to get a specific Topic."
        Write-Wrap "e.g. 'm op' will display 9 matches. Refine to 'm compar' to open 'about_Comparison_Operators'"
        Write-Wrap "By default, will open files as 'Get-Help <Help-File> -detail | more' to show switch details and examples."
        Write-Wrap "Note that when '| more' is used, any key will go forward one page, but <right-click> will escape the '| more' and show all remaining text."
        Write-Host ""
        if (! ( ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) ) ) { "No elevated privileges, so Help System updates cannot run.`n" ; break }
        Write-Host ""
        Write-Host ""
        Write-Host ""
        ##########
        # Admin only section - Everything after here must be Admin, as if not, will break out in the above test
        ##########
        $updatecached = read-host "Run 'Update-Help' and update all lists of help files in user temp folder (default is 'n') [y/n]? "
        if ($updatecached -eq 'y') {
            # Update-Help -EA silent   # Use erroraction silent here as it's very common for some of the modules to fail to update, just ignore that.
            $updatefile = "$($env:TEMP)\ps_Update-Help-$($PSVersionTable.PSVersion.ToString()).flag"
            [int]$helpolderthan = 20
            [datetime]$dateinpast = (Get-Date).AddDays(-$helpolderthan)

            if (Test-Path $updatefile) { [datetime]$updatetime = (Get-Item $updatefile).LastWriteTime }
            
            Write-Host ""
            Write-Host ""
            Write-Host "n========================================" -ForegroundColor Green
            Write-Host "Update Help Files if more than $helpolderthan days old." -F Yellow -B Black
            Write-Host "Checking PowerShell Help definitions ..." -F Yellow -B Black
            Write-Host "" 
            Write-Host "Note: This section will only show if you are running as Administrator." 
            Write-Host "========================================" -ForegroundColor Green
            Write-Host ""
            if ($PSVersionTable.PSVersion.Major -eq 2) {
                Write-Host "Update-Help Cmdlet does not exist on PowerShell v2" -ForegroundColor Red
                Write-Host "Skipping help definitions update ..." -ForegroundColor Red
            } else {
                if (Test-Path $updatefile) {
                    "Current Date minus $helpolderthan         : $((Get-Date).AddDays(-$helpolderthan))"   # Note that this has a "-" so this is back 20 days from the current day
                    "Date on help update flag file : $updatetime"
                    if ($updatetime -lt $dateinpast) {
                        Write-Host "Help files only update if more than $($helpolderthan.ToString()) days old. A flag file is" -ForegroundColor Green
                        Write-Host "kept in the user Temp folder timestamped at last update time to check this ..." -ForegroundColor Green
                        Write-Host "The flag file is more than days old, so Help file definitions will update ..." -ForegroundColor Green
                        (Get-ChildItem $updatefile).LastWriteTime = Get-Date   # touch the flat file to today's date
                        Update-Help -ErrorAction SilentlyContinue
                        # Run this with -EA silent due to the various errors that always happen as a result of bad manifests, which will be similar to:
                        # Note that this will not suppress all error messages! so need to add | Out-Null also
                        #   update-help : Failed to update Help for the module(s) 'AppvClient, ConfigDefender, Defender, HgsClient, HgsDiagnostics, HostNetworkingService,
                        #   Microsoft.PowerShell.ODataUtils, Microsoft.PowerShell.Operation.Validation, UEV, Whea, WindowsDeveloperLicense' with UI culture(s) {en-GB} : Unable to
                        #   connect to Help content. The server on which Help content is stored might not be available. Verify that the server is available, or wait until the server is
                        #   back online, and then try the command again.
                    } else {
                        Write-Wrap "The flag file used to determine if help definitions should be updated is less than $($helpolderthan.ToString()) days old so help file will not be updated ..."
                    }
                } else {
                    Write-Host "No help definitions checkfile found, creating a new checkfile and Update-Help will run ..."
                    New-Item -ItemType File $updatefile
                    Update-Help -ErrorAction SilentlyContinue
                }
            }
            Write-Host ""
            ##########
            # ToDo: check for duplicate help files and just delete the oldest ones!
            if (-not (Test-Path $about_name)) { ($about | select Name | sort Name -Unique | Out-String) -replace "[Aa]bout_", "" > $about_name }
            if (-not (Test-Path $about_namecsv)) { (($about | select Name | sort Name -Unique | % { $_.Name + "," } | Out-String) -replace "[Aa]bout_", "" -replace "`r`n", " ").trim(", ") > $about_namecsv }
            if (-not (Test-Path $about_synopsis)) { ($about | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } | Out-String) -replace "[Aa]bout_", "" > $about_synopsis }
            if (-not (Test-Path $help_name)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name | sort Name -Unique > $help_name }
            if (-not (Test-Path $help_synopsis)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } > $help_synopsis }
            # $job_about_name = Start-Job { if (-not (Test-Path $using:about_name)) { ($using:about | select Name | sort Name -Unique | Out-String) -replace "[Aa]bout_", "" > $using:about_name } }
            # $job_about_namecsv = Start-Job { if (-not (Test-Path $using:about_namecsv)) { (($using:about | select Name | sort Name -Unique | % { $_.Name + "," } | Out-String) -replace "[Aa]bout_", "" -replace "`r`n", " ").trim(", ") > $using:about_namecsv } }
            # $job_about_synopsis = Start-Job { if (-not (Test-Path $using:about_synopsis)) { ($using:about | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } | Out-String) -replace "[Aa]bout_", "" > $using:about_synopsis } }
            # $job_help_name = Start-Job { if (-not (Test-Path $using:help_name)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name | sort Name -Unique > $using:help_name } }
            # $job_help_synopsis = Start-Job { if (-not (Test-Path $using:help_synopsis)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } > $using:help_synopsis } }
            # Wait-Job $job_about_name, $job_about_namecsv, $job_about_synopsis, $job_help_name, $job_help_synopsis 
        }
        ""
    } else {

        if ((cat $about_name) -match [regex]::Escape($cmd)) {   # check for matches in here ... [regex]::Escape($x) forces a simplematch!
            $matchlist = (cat $about_name) -match [regex]::Escape($cmd)        # do not sort this as it breaks some later code
            ##########
            #
            # This section handles about_* output
            #
            ##########
            $n = ($matchlist).count
            if ($n -eq 0) {
                # could be a non-about, check against other sources?
            }

            if ($n -eq 1) {
                $matchname = ($matchlist[0]).ToString().Trim()
                if ($Examples -eq $true) {
                    Get-Help $matchname -Examples | more
                }
                # about_* output does not use -Examples, -Synopsis, -Syntax (!)
                #
                # elseif ($Synopsis -eq $true) {
                #     Write-Host "`nSYNOPSIS for $matchname   " -F Green -NoNewline ; Write-Host "(Get-Help $matchname).Synopsis`n" -F Cyan
                #     $arrsynopsis = ((Get-Help $cmd).Synopsis).TrimStart("").Split("`n")  # Trim empty first line then split by line breaks
                #     foreach ($i in $arrsynopsis) { Write-Wrap $i }   # Wrap lines properly to console width
                #     Write-Host ""
                # }
                # elseif ($Syntax -eq $true) {
                #     Write-Host "`nSYNTAX for $matchname   " -F Green -NoNewline ; Write-Host "Get-Command $matchname -Syntax" -F Cyan
                #     $arrsyntax = (Get-Command $cmd -syntax).TrimStart("").Split("`n")    # Often synopsis=syntax for function so use Compare-Object
                #     foreach ($i in $arrsyntax) { Write-Wrap $i }     # Wrap lines properly to console width
                #     Write-Host ""
                # }
                else {
                    Get-Help $matchname -Detailed | more
                }

            } else {
                $x = @()
                $show = $null
                foreach ($i in $matchlist) {
                    # Need to collect them first and then remove duplicates
                    $search = $($i.ToString().Trim())
                    $x += $out_synopsis -match [regex]::Escape($search)
                    if ($cmd -imatch "$($i.ToString().Trim())") { $show = $cmd }   # reversing search looking for exact match
                }
                $x = $x | sort -Unique
                echo ""
                foreach ($i in $x) { Write-Wrap "$i" ; "" }

                if ($show -eq $null) {
                    if ($foundmethod -eq $true) { Write-Host "`n------------`n"}
                    ""
                    Write-Wrap "The search term '$cmd' was found in the $n about_ Topics above. Refine the search term to open a specific about_ Topic file."
                    ""
                }
                else {
                    ""
                    Write-Wrap "The search term '$cmd' was found in the $n about_ Topics above"
                    Write-Wrap "but also has an exact match  that will now be opened."
                    ""
                    Read-Host "Press any key to open the matched Topic"
                    ""
                    ""
                    Get-Help "about_$show" | more
                }

                # This is all broken, need to check the original regex
            }

        } else {

            echo "`nNo about_ Topics found for '$cmd'. Searching among help / alias entries ..."
            ##########
            #
            # This section handles help file (non- about_*) output
            #
            ##########

            # First have to check if $cmd is an alias. If it is, get the full command and use that
            $aliasdef = (get-alias $cmd -EA Silent).Definition

            if ($null -ne $aliasdef) {
                # In this case, the base command was found, so deal with it like that
                if ($cmd -eq '?') { $cmd = '`?' }   # To deal correctly with the wildcard '?'
                "`n'$((Get-Alias $cmd).Name)' is an alias of '$((Get-Alias $cmd).ReferencedCommand)'"

                if ($Examples -eq $true) {
                    Get-Help $aliasdef -Examples | more
                }
                elseif ($Synopsis -eq $true) {
                    Write-Host "`nSYNOPSIS for $aliasdef   " -F Green -NoNewline ; Write-Host "(Get-Help $aliasdef).Synopsis" -F Cyan
                    $arrsynopsis = ((Get-Help $aliasdef).Synopsis).TrimStart("").Split("`n")  # Trim empty first line then split by line breaks
                    foreach ($i in $arrsynopsis) { Write-Wrap $i }   # Wrap lines properly to console width
                    Write-Host ""
                }
                elseif ($Syntax -eq $true) {
                    Write-Host "`nSYNTAX for $aliasdef   " -F Green -NoNewline ; Write-Host "Get-Command $aliasdef -Syntax" -F Cyan
                    $arrsyntax = (Get-Command $aliasdef -syntax).TrimStart("").Split("`n")    # Often synopsis=syntax for function so use Compare-Object
                    foreach ($i in $arrsyntax) { Write-Wrap $i }     # Wrap lines properly to console width
                    Write-Host ""
                }
                else {
                    Get-Help $aliasdef -Detailed | more
                }
                break   # Stop as no need to continue, we found an alias root command got required output
            }

            $out_name = cat $help_name
            if (! (Test-Path $help_synopsis)) { if (-not (Test-Path $help_synopsis)) { get-help * | ? { $_.Name -NotMatch '^about' } | select Name,Synopsis | sort Name -Unique | % { $_.Name + " :: " + $_.Synopsis } > $help_synopsis } }
            $out_synopsis = cat $help_synopsis
            
            # check for matches in $help_name
            if ((cat $help_name) -match [regex]::Escape($cmd)) {   
                $matchlist = (cat $help_name) -match [regex]::Escape($cmd)   # sort Name breaks this
                $n = ($matchlist).count
                if ($n -eq 0) {
                    # could be a non-about, check against other sources
                }
                if ($n -eq 1) {
                    $matchname = ($matchlist[0]).ToString().Trim()
                    if ($Examples -eq $true) {
                        Get-Help $matchname -Examples | more
                    }
                    elseif ($Synopsis -eq $true) {
                        Write-Host "`nSYNOPSIS for $matchname   " -F Green -NoNewline ; Write-Host "(Get-Help $matchname).Synopsis" -F Cyan
                        $arrsynopsis = ((Get-Help $cmd).Synopsis).TrimStart("").Split("`n")  # Trim empty first line then split by line breaks
                        foreach ($i in $arrsynopsis) { Write-Wrap $i }   # Wrap lines properly to console width
                        Write-Host ""
                    }
                    elseif ($Syntax -eq $true) {
                        Write-Host "`nSYNTAX for $matchname   " -F Green -NoNewline ; Write-Host "Get-Command $matchname -Syntax" -F Cyan
                        $arrsyntax = (Get-Command $cmd -syntax).TrimStart("").Split("`n")    # Often synopsis=syntax for function so use Compare-Object
                        foreach ($i in $arrsyntax) { Write-Wrap $i }     # Wrap lines properly to console width
                        Write-Host ""
                    }
                        else {
                        Get-Help $matchname -Detailed | more
                    }
                } 
                else {
                    $x = @()
                    foreach ($i in $matchlist) {
                        # Need to collect them first and then remove duplicates
                        $x += $out_synopsis -match [regex]::Escape($($i.ToString().Trim()))   # use [regex]::Escape() to escape all special characters!
                    }
                    $x = $x | sort -Unique
                    echo ""
                    foreach ($i in $x) { Write-Wrap "$i" }
                }
            } else {
                echo "Nothing was found from 'help *' ...`nLast attempt is to just try:`n   help $cmd | more`n`n"
                help $cmd | more
            }
        }
    }
}

# Man(uals) S(earch). ToDo, fully search for <$search> in files matching $help if that is provided (or all help files if not, could be slow that way)
function Search-ManPagesWIP ($search, $help) {
    "ToDo / WIP ... Should take a `$search and use that to search search all help files that have `$help in the name."
    "`$help should be optional, if not there, search *all* help files (maybe cache all output to $env:TEMP for future use)"
    "ms on its own could update the help file cache?"
}
# $arrdescription = (get-help $cmd).Description.Text.split("`n")
# foreach ($i in $arrdescription) { Write-Wrap $i }
# foreach($line in Get-Content .\file.txt)   # This is the wrong way to do it, not pipeline aware, loads entire file into memory
# Get-Content $about_name | % { echo $_ }   # pipeline aware, will load line by line through the pipeline
# Get-Content $about_name | % { echo $_ }
# $name_synopsis
# $arrdescription = (get-help $cmd).Description.Text.split("`n")
# foreach ($i in $arrdescription) { Write-Wrap $i }
# Man Search
# ms : search multiple help files based on help-name-string / search-string
# test for the .txt files
#     Get-Help about_* > $help_about ; cat $help_about | more ; break
#     Write-Host "about_ Topic Name :: Topic Synopsis`n"
# } else {
#     # if something is given, a) assume about_, then test for full name
#     try { Get-Help about_$($cmd) | more ; break }
#     catch { echo "failed 1"}
# }
#     Write-Host "about_ Topic Name :: Topic Synopsis`n"
#     # Just list topics Name and Sysnopsis fields ...
#     (get-help about_* | select Name,Synopsis | % { $_.Name + " :: " + $_.Synopsis } | Out-String).replace("about_", "")
#     # Then just list about_ topics as a list ...
#     Write-Host "Quick list of all current topics`n"
#     (get-help about_* | select Name | % { $_.Name + "," } | Out-String).replace("about_", "").replace("`r`n", " ").trim(", ")
# }

# Might have a separate function ha-search => download all about_ to $env:temp
# Build searchable dump of all about_* topics, took ~1 min to run for 157 topics
# Search all about_* for a keyword
# Get-Help about_* | select Name | % { Get-Help $_.Name > "$($env:TEMP)\$($_.Name).txt" }
# foreach ($i in "$($env:TEMP)\about_*.txt" }
# function ha-search($search) { foreach ($i in "$($env:TEMP)\about_*.txt") { echo "$($i):"  ; cat $i | sls $search } }



# Find definitions for any Cmdlet, Function, Alias, External Script, Application
function what {   
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ArgumentCompleter({ [Management.Automation.CompletionResult]::Command })]
        $cmd,
        [switch]$Examples
    )
    # Previously declared $cmd as [string]$cmd but this was wrong as cannot then handle arrays or anything else

    function Write-Wrap {
        [CmdletBinding()]Param( [parameter(Mandatory=1, ValueFromPipeline=1, ValueFromPipelineByPropertyName=1)] [Object[]]$chunk )
        $Lines = @()
        foreach ($line in $chunk) {
            $str = ''; $counter = 0
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

    $deferr = 0; $type = ""
    try { $type = ((gcm $cmd -EA silent).CommandType); if ($null -eq $type) { $deferr = 1 } } catch { $deferr = 1 }

    if ($deferr -eq 1) {
        if ($cmd -eq $null) { Write-Host "Object is `$null" ; return } 
        Write-Host "`$object | ConvertTo-Json:" -F Cyan
        $cmd | ConvertTo-Json
        ""
        Write-Host "(`$object).GetType()" -F Cyan -NoNewline ; Write-Host " (Below is: BaseType, Name, IsPublic, IsSerial, Module)"
        ($cmd).GetType() | % { "$($_.BaseType), $($_.Name), $($_.IsPublic), $($_.IsSerializable), $($_.Module)" }
        ""
        Write-Host "`$object | Get-Member -Force" -F Cyan
        $m = "" ; $cm = "" ; $sm = ""; $p = "" ; $ap = "" ; $cp = "" ; $np = "" ; $pp = "" ; $sp = "" ; $ms = ""
        $msum = 0 ; $cmsum = 0 ; $smsum = 0 ; $psum = 0 ; $cpsum = 0 ; $apsum = 0 ; $spsum = 0 ; $ppsum = 0 ; $npsum = 0 ; $spsum = 0 ; $mssum = 0
        $($cmd | Get-Member -Force) | % {
            if ($_.MemberType -eq "Method") { if(!($m -like "*$($_.Name),*")) { $m += "$($_.Name), " ; $msum++ } }
            if ($_.MemberType -eq "CodeMethod") { if(!($cm -like "*$($_.Name),*")) { $cm += "$($_.Name), " ; $cmsum++ } }
            if ($_.MemberType -eq "ScriptMethod") { if(!($sm -like "*$($_.Name),*")) { $sm += "$($_.Name), " ; $smsum++ } }
            if ($_.MemberType -eq "Property") { if(!($p -like "*$($_.Name),*")) { $p += "$($_.Name), " ; $psum++ } }
            if ($_.MemberType -eq "AliasProperty") { if(!($ap -like "*$($_.Name),*")) { $ap += "$($_.Name), " ; $apsum++ } }
            if ($_.MemberType -eq "CodeProperty") { if(!($cp -like "*$($_.Name),*")) { $cp += "$($_.Name), " ; $cpsum++ } }
            if ($_.MemberType -eq "NoteProperty") { if(!($np -like "*$($_.Name),*")) { $np += "$($_.Name), " ; $npsum++ } }
            if ($_.MemberType -eq "ParameterizedProperty") { if(!($pp -like "*$($_.Name),*")) { $pp += "$($_.Name), " ; $ppsum++} }
            if ($_.MemberType -eq "ScriptProperty") { if(!($sp -like "*$($_.Name),*")) { $sp += "$($_.Name), " ; $npsum++ } }
            if ($_.MemberType -eq "MemberSet") { if(!($ms -like "*$($_.Name),*")) { $ms += "$($_.Name), " ; $mssum++ } }
            # AliasProperty, CodeMethod, CodeProperty, Method, NoteProperty, ParameterizedProperty, Property, ScriptMethod, ScriptProperty
            # All, Methods, MemberSet, Properties, PropertySet
        }
        if($msum -ne 0) { Write-Wrap ":: Method [$msum] => $($m.TrimEnd(", "))" }
        if($msum -ne 0) { Write-Wrap ":: CodeMethod [$cmsum] => $($cm.TrimEnd(", "))" }
        if($msum -ne 0) { Write-Wrap ":: ScriptMethod [$smsum] => $($sm.TrimEnd(", "))" }
        if($psum -ne 0) { Write-Wrap ":: Property [$psum] => $($p.TrimEnd(", "))" }
        if($npsum -ne 0) { Write-Wrap ":: AliasProperty [$apsum] => $($ap.TrimEnd(", "))" }
        if($npsum -ne 0) { Write-Wrap ":: CodeProperty [$cpsum] => $($cp.TrimEnd(", "))" }
        if($npsum -ne 0) { Write-Wrap ":: NoteProperty [$npsum] => $($np.TrimEnd(", "))" }
        if($ppsum -ne 0) { Write-Wrap ":: ParameterizedProperty [$ppsum] => $($pp.TrimEnd(", "))" }
        if($spsum -ne 0) { Write-Wrap ":: ScriptProperty [$spsum] => $($sp.TrimEnd(", "))" }
        if($mssum -ne 0) { Write-Wrap ":: ScriptProperty [$mssum] => $($ms.TrimEnd(", "))" }
        ""
        Write-Host "`$object | Measure-Object" -F Cyan
        $cmd | Measure-Object | % { "Count [$($_.Count)], Average [$($_.Average)], Sum [$($_.Sum)], Maximum [$($_.Maximum)], Minimum [$($_.Minimum)], Property [$($_.Property)]" }
    }

    if ($deferr -eq 0) {

        if ($cmd -like '*`**') { Get-Command $cmd ; break }   # If $cmd contains a *, then just check for commands, don't find definitions
   
        if ($type -eq 'Cmdlet') {
            Write-Host "`n'$cmd' is a Cmdlet:`n" -F Green
            Write-Host "SYNOPSIS, DESCRIPTION, SYNTAX for '$cmd'.   " -F Green
            Write-Host "------------"
            Write-Host ""
            Write-Host "(Get-Help $cmd).Synopsis" -F Cyan 
            Write-Host "$((Get-Help $cmd).Synopsis)"
            Write-Host ""
            Write-Host "(Get-Help $cmd).Description.Text" -F Cyan
            try {
                $arrdescription = (Get-Help $cmd).Description.Text.split("`n")
                foreach ($i in $arrdescription) { Write-Wrap $i }
            } catch { "Could not resolve description for $cmd" }
            Write-Host ""
            Write-Host "(Get-Command $cmd -Syntax)" -F Cyan
            $arrsyntax = (Get-Command $cmd -syntax).TrimStart("").Split("`n")  # Trim empty first line then split by line breaks
            foreach ($i in $arrsyntax) { Write-Wrap $i }   # Wrap lines properly to console width
            Get-Alias -definition $cmd -EA silent          # Show all defined aliases
            Write-Host "`nThis Cmdlet is in the '$((Get-Command -type cmdlet $cmd).Source)' Module." -F Green
            Write-Host ""
            Write-Host ""
        }
        elseif ($type -eq 'Alias') {
            Write-Host "`n'$cmd' is an Alias.  " -F Green -NoNewLine ; Write-Host "This Alias is in the '$((get-command -type alias $cmd).ModuleName).' Module"
            Write-Host ""
            Write-Host "Get-Alias '$cmd'   *or*    cat alias:\$cmd" -F Cyan
            $aliasdef = $(cat alias:\$cmd)   # Write-Host "$(cat alias:\$cmd)"   # "$((Get-Alias $cmd -EA silent).definition)"
            if ($cmd -eq '?') { $cmd = '`?' }   # To deal correctly with the wildcard '?'
            $cmdref = (Get-Alias $cmd).ReferencedCommand
            if ($null -eq $cmdref) {
                "`n'$((Get-Alias $cmd).Name)' is an alias of '$aliasdef', but '$aliasdef' is not a defined command."
                "cat alias:\$cmd                     =>  $aliasdef"
                "(Get-Alias $cmd).ReferencedCommand  =>  `$null"
            } else {
                "`n'$((Get-Alias $cmd).Name)' is an alias of '$cmdref'"   # $((Get-Alias $cmd).ReferencedCommand)
                $fulldef = (Get-Alias $cmd -EA silent).definition   # Rerun def but using the full cmdlet or function name.
                def $fulldef
                if ($Examples -eq $true) { $null = Read-Host 'Press any key to view command examples' ; get-help $fulldef -examples }
            }
        }
        elseif ($type -eq 'Function') {
            Write-Host "`n'$cmd' is a Function.  " -F Green -NoNewline
            Write-Host "`ncat function:\$cmd   (show contents of function)`n" -F Cyan
            if ($bat = Get-Command bat -ErrorAction Ignore) {
                (Get-Content function:$cmd) | & $bat -p -l powershell
            } else {
                cat function:\$cmd ; Write-Host ""
            }
            Write-Host "cat function:\$cmd`n" -F Cyan
            Write-Host ""
            Write-Host "SYNOPSIS, SYNTAX for '$cmd'.   " -F Green
            Write-Host "------------"
            $arrsynopsis = ((Get-Help $cmd).Synopsis).TrimStart("").Split("`n")  # Trim empty first line then split by line breaks
            $arrsyntax = (Get-Command $cmd -syntax).TrimStart("").Split("`n")    # Often synopsis=syntax for function so use Compare-Object
            if ($null -eq $(Compare-Object $arrsynopsis $arrsyntax -SyncWindow 0)) { 
                Write-Host "'(Get-Help $cmd).Synopsis'" -F Cyan -N
                Write-Host " and " -N
                Write-Host "'Get-Command $cmd -Syntax'" -F Cyan -N
                Write-Host " have the same output for this function:`n"
                foreach ($i in $arrsynopsis) { Write-Wrap $i }   # Wrap lines properly to console width
            } else { 
                Write-Host "(Get-Help $cmd).Synopsis" -F Cyan
                foreach ($i in $arrsynopsis) { Write-Wrap $i }   # Wrap lines properly to console width
                Write-Host ""
                Write-Host "Get-Command $cmd -Syntax" -F Cyan
                foreach ($i in $arrsyntax) { Write-Wrap $i }     # Wrap lines properly to console width
            }
            Write-Host "The '$cmd' Function is in the '$((get-command -type function $cmd).Source)' Module." -F Green
            Write-Host ""
            if ($Examples -eq $true) { $null = Read-Host "Press any key to view command examples" ; get-help $cmd -examples }
            Write-Host ""
        }
        elseif ($type -eq 'ExternalScript') {   # For .ps1 scripts in current location or on the path
            $x = gcm $cmd
            Write-Host "`n'$cmd' is an ExternalScript (i.e. a .ps1 file in current location or on the path)." -F Green
            Write-Host "`n$($x.Path)`n" -F Green
            Write-Host "`n$($x.ScriptContents)"
            Write-Host ""
            if ($Examples -eq $true) { $null = Read-Host "Press any key to view command examples" ; get-help $cmd -Examples }
            elseif ($Synopsis -eq $true) { $null = Read-Host "Press any key to view command synopsis" ; (get-help $cmd).Synopsis }
            elseif ($Syntax -eq $true) { $null = Read-Host "Press any key to view command syntax" ; Get-Command $cmd -Syntax }
            Write-Host ""
        }
        elseif ($type -eq 'Application') {      # For .exe etc on path, or could also be a .cmd / .bat etc
            Write-Host "`n'$cmd' was found. It is an Application (i.e. a .exe or similar located on the path)." -F Green
            where.exe $cmd
            # if .cmd / .bat, then show it with correct -l setting
            # if ($bat = Get-Command bat -ErrorAction Ignore) {
            #     (Get-Content function:$cmd) | & $bat -pp -l powershell
            # } else {
            #     offer the /? option otherwise
            # }
            Write-Host ""
            Read-Host "Press any key to open cmd.exe and try '$cmd /?'" ; cmd.exe /c $cmd /? | more
            Write-Host ""
        }
    } elseif ($null -ne (get-module -ListAvailable -Name $cmd -EA Silent)) {
        # https://stackoverflow.com/questions/28740320/how-do-i-check-if-a-powershell-module-is-installed
        ""
        (get-module $cmd).path
        (get-module $cmd).ExportedFunctions
        "ExportedCommands (also note: get-command -Module $cmd)"
        (get-module custom-tools).ExportedCommands
        ""
        echo "get-module $cmd | get-member  # Just show the members"
        echo "get-module $cmd | fl *        # Show the contents of every member"
    }
    else {
        if ($cmd.length -eq 0) { "`n'$cmd': No command definition found. The command may require to be surround by ' or `"`nif it contains special characters (such as 'def `"&`"').`n" }
        else { "`nInput is not a command, so no command definition search.`n" }
    }
}

Set-Alias def what   # Generally better to use "what" (def is used in various other languages), but I'll keep this aliased in PowerShell for convenience

# Old method, test if an error happens. Above is much better, but this could be useful elsewhere
# $defx = 0
# $error.clear()
# try { get-alias $cmd -EA silent | Out-Null } catch { $defx += 1 }
# if (!$error) { Write-Host "`n'$cmd' was found. It is an Alias  " -F Green -NoNewLine ; Write-Host "Get-Alias $cmd" -F Cyan ; Get-Alias $cmd }
# $error.clear()
# try { cat function:\$cmd -EA silent | Out-Null } catch { $defx += 1 }
# if (!$error) { Write-Host "`n'$cmd' was found. It is a Function:  " -F Green -NoNewLine ; Write-Host "cat function:\$cmd`n" -F Cyan ; cat function:\$cmd ; Write-Host "" }
# if ($defx -eq 2) { echo "No definition was found for '$cmd'`n"}

function CookieKiller ($search) {
    # Note shortcut to Chrome cookies:  Ctrl-Shift-Delete  /  chrome://settings/clearBrowserData  /  chrome://settings/siteData
    # C:\Users\<your_username>\AppData\Local\Google\Chrome\User Data\Default\
    # Idea is to delete all cookies that match $search
    # This might not be possible, since this stuff is all controlled internally within Chrome.
    # Cookie Keeper might be best option, set domain rules for cookies you want to keep, all other cookies are deleted.
    # https://chrome.google.com/webstore/detail/cookie-keeper/hkdjopjjoogbicnmbcenniplfnmcnhof
    # An alternative might be to use AutoHotkey, natigate to chrome://settings/content/cookies, etc
    # get-childitem "C:\Users\testuser1\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -include * -recurse -force | ? {$_.name -match "videoplayback"} | Remove-item -force -recurse 
    # The above tested on the win7 machine. Change the location and matching criteria.
    # The below is the single command to remove the browser history: RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
    # PowerShell remove cookies (but maybe IE or Edge only?): http://csoposh.blogspot.com/2011/11/delete-cookies.html
    # https://vworld.nl/?p=3881
    # Firefox Cookies API: https://developer.mozilla.org/en-US/docs/Mozilla/Add-ons/WebExtensions/Work_with_the_Cookies_API
    # https://github.com/PoE-TradeMacro/POE-TradeMacro/issues/173
    # Test Browser Performance : https://helgeklein.com/blog/2018/12/powershell-script-test-chrome-firefox-ie-browser-performance/
}

# Need to update this so that it will purge the choco lines from the profile...
function Enable-Choco {
    # Removed from $Profile as 1.5 s load time, can load this whenever required with this function 
    $ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
    if(Test-Path $ChocolateyProfile) {
        Import-Module "$ChocolateyProfile"
    }
}

function IfExistSkipCommand ($toCheck, $toRun) {
    if (Test-Path($toCheck)) {
        Write-Host "Item exists        : $toCheck" -ForegroundColor Green
        Write-Host "Will skip installer: $toRun`n" -ForegroundColor Cyan
    } else {
        Write-Host "Item does not exist: $toCheck" -ForegroundColor Green
        Write-Host "Will run installer : $toRun`n" -ForegroundColor Cyan
        Invoke-Expression $toRun

    }
}

function Install-PowershellCore {
    Write-Host "Get latest PowerShell Core"
    # PowerShell Core, get latest version
    choco upgrade -y PowerShell-Core
}

function Install-Firefox {

    Write-Wrap "When you continue, Firefox processes will be killed and redundant copies of Firefox will be uninstalled from the Users AppData folder and from C:\Program Files (x86). The 64-bit Firefox only will be left, or will be installed if not present."
    ""
    Write-Wrap "- will taskkill.exe /f /im firefox.exe (but only if already running)"
    Write-Wrap "- will taskkill.exe /f /im firefox.exe (but only if already running)"
    Write-Wrap "- will taskkill.exe /f /im firefox.exe (but only if already running)"
    pause

    if (Test-Path "$env:USERPROFILE\AppData\Local\Mozilla Firefox\uninstall\helper.exe") {
        taskkill.exe /f /im firefox.exe
        $setup = "$env:USERPROFILE\AppData\Local\Mozilla Firefox\uninstall\helper.exe"
        $uninst = Start-Process $setup -PassThru -ArgumentList "/s" -wait
        $uninst.WaitForExit()
    }

    if (Test-Path "${env:ProgramFiles(x86)}\Mozilla Firefox\uninstall\helper.exe") {
        taskkill.exe /f /im firefox.exe
        $setup = "${env:ProgramFiles(x86)}\Mozilla Firefox\uninstall\helper.exe"
        $uninst = Start-Process $setup -PassThru -ArgumentList "/s" -wait
        $uninst.WaitForExit()
    }

    # if (Test-Path ($env:ProgramFiles + "\Mozilla Firefox\uninstall\helper.exe") ) {
    #     taskkill.exe /f /im firefox.exe
    #     $setup = $env:ProgramFiles + "\Mozilla Firefox\uninstall\helper.exe"
    #     $args = " /s"
    #     $uninst = Start-Process $setup -PassThru -ArgumentList $args -wait
    #     $uninst.WaitForExit()
    # }

    # Detection

    # if (-not (Test-Path ($env:ProgramFiles + "\Mozilla Firefox\uninstall\helper.exe") ) -and 
    #     -not (Test-Path ($env:USERPROFILE + "\AppData\Local\Mozilla Firefox\uninstall\helper.exe") ) -and
    #     -not (Test-Path (${env:ProgramFiles(x86)} + "\Mozilla Firefox\uninstall\helper.exe") ) )

    if (-not (Test-Path "$env:ProgramFiles\Mozilla Firefox\uninstall\helper.exe")) {
        Write-Host "Downloading Firefox 64-bit and run installation (not yet implemented)..."
        pause
    }
}

function LoginToWebSite ($siteurl) {
    # https://stackoverflow.com/questions/40624990/auto-login-to-a-website-using-powershell
    # https://social.technet.microsoft.com/Forums/Lync/en-US/be3afe83-4a7e-48a0-b2e7-95fd081a7571/login-to-website-using-powershell?forum=winserverpowershell
    # https://social.technet.microsoft.com/Forums/ie/en-US/9f214ed6-66fc-43d1-b775-d841bddcbcfa/cannot-get-authentication-cookie-from-web-server-with-invokewebrequest?forum=ITCG
}

# Force redownload of Custom-Tools to C:\Program Files\WindowsPowerShell\Modules\Custom-Tools
# and force re-import.
function Remove-Toolkit {

    # Rollback changes in BeginSystemConfig.ps1 and disable ProfileExtensions.ps1 / Custom-Tools.psm1
    # 1. Find project files, either locally, or from internet
    # 2. Parse BeginSystemConfig.ps1
    #    Look for configured Modules and try to uninstall and remove
    #    Look for configured Scripts and remove
    # 3. Remove the reference to ProfileExtensions from $Profile
    # 4. Delete all ProfileExtensions from the $Profile folder.

    $ProjectRoot = $null
    if (Test-Path "$HomeFix\Gist") { $ProjectRoot = "$HomeFix\Gist" }
    if (Test-Path "D:\Gist") { $ProjectRoot = "D:\Gist" }
    if (Test-Path "D:\0 Cloud\OneDrive\Gist") { $ProjectRoot = "D:\0 Cloud\OneDrive\Gist" }
    echo $ProjectRoot
    if ($ProjectRoot -eq $null) {
        # get from internet
        $UrlConfig = 'https://gist.github.com/roysubs/61ef677591f22927afadc9ef2b657cd9/raw'
        $UrlProfileExtensions = 'https://gist.github.com/roysubs/c37470c98c56214f09f0740fcb21ec4f/raw'
        $UrlCustomTools = 'https://gist.github.com/roysubs/5c6a16ea0964cf6d8c1f9eed7103aec8/raw'
        try { (New-Object System.Net.WebClient).DownloadString("$UrlConfig") | Out-File "$($env:TEMP)\BeginSystemConfig.ps1" ; echo "Downloaded latest system config script from internet ..."}
        catch { Write-Host "BeginSystemConfig.ps1 failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
        try { (New-Object System.Net.WebClient).DownloadString("$UrlProfileExtensions") | Out-File "$($env:TEMP)\ProfileExtensions.ps1" ; echo "Downloaded latest profile extensions from internet ..."}
        catch { Write-Host "ProfileExtensions.ps1 failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
        try { (New-Object System.Net.WebClient).DownloadString("$UrlCustomTools") | Out-File "$($env:TEMP)\Custom-Tools.psm1" ; echo "Downloaded latest Custom-Tools.psm1 from internet ..."}
        catch { Write-Host "Custom-Tools.psm1 failed to download! Check internet/TLS settings before retrying." -ForegroundColor Red }
    }
    else {
        # get locally
        if ($ProjectRoot -ne "$($env:TEMP)") {   # When elevating to admin, the called script is in TEMP, so skip copying as will be to same location!
            if (Test-Path "$($ProjectRoot)\BeginSystemConfig.ps1") { Copy-Item "$($ProjectRoot)\BeginSystemConfig.ps1" "$($env:TEMP)\BeginSystemConfig.ps1" -Force }
            if (Test-Path "$($ProjectRoot)\ProfileExtensions.ps1") { Copy-Item "$($ProjectRoot)\ProfileExtensions.ps1" "$($env:TEMP)\ProfileExtensions.ps1" -Force }
            if (Test-Path "$($ProjectRoot)\Custom-Tools.psm1")     { Copy-Item "$($ProjectRoot)\Custom-Tools.psm1" "$($env:TEMP)\Custom-Tools.psm1" -Force }
        }
    }
    pause
    # operate on the files, rollback modules etc
    Get-Content -Path "$($env:TEMP)\BeginSystemConfig.ps1" | Select-String -Pattern "Install-Module.+" | ForEach-Object {
        [Regex]::Matches($_, "{ Install Module ([a-z.-]+)","IgnoreCase").Groups[1].Value
    } | Where-Object { $_ } | Sort-Object
    # get-content .\BeginSystemConfig.ps1 | Select-String -Pattern "{ Install-Module.+"

    # cleanup temp files
    Remove-Item "$($env:TEMP)\BeginSystemConfig.ps1" -Force
    Remove-Item "$($env:TEMP)\ProfileExtensions.ps1" -Force
    Remove-Item "$($env:TEMP)\Custom-Tools.psm1" -Force
}



function Disable-BitLockerEncryption {
    ""

    function Test-Administrator {
        $user = [Security.Principal.WindowsIdentity]::GetCurrent();
        (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    if (Test-Administrator -eq $false) {
        "Cannot continue as must be Administrator to change required rgistry/policy settings:"
        "Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker -Name PreventDeviceEncryption -Value 1   # Set to 0 to Enable Encryption"
        "Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name RDVDenyWriteAccess -Value 0   # Set to 1 to turn on the Policy"
    }

    "Note that encyption will re-enable in time when the policy is reapplied"
    "Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker -Name PreventDeviceEncryption -Value 1   # Set to 0 to Enable Encryption"
    Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker -Name PreventDeviceEncryption -Value 1
    ""
    "# https://jessehouwing.net/windows-bitlocker-bypass-temporarily/"
    "Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name RDVDenyWriteAccess -Value 0   # Set to 1 to turn on the Policy"
    Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name RDVDenyWriteAccess -Value 0
    ""
    "After disabling BitLocker encryption, have to unmount the drive and then remount it."
    "Normally that is done by unplugging and replugging it in, automating this is a bit tricky in Windows."
    "To unmount a volume from the command line:"
    "   mountvol D: /p     # type 'mountvol' to see options"
    "Some sites say that you can then re-mount that drive using the GUID shown in 'mountvol'"
    "   mountvol D: \\?\Volume{24f355f9-5bce-4536-b265-4f0236458071}\"
    ""
    "I cannot get the above to work but the alternative is to run 'diskmgmt.msc' and then assign"
    "a letter to the drive which gets around the need to physically unplug and then re-plug in the drive."
    "https://superuser.com/questions/295913/how-to-mount-and-unmount-hard-drives-under-windows-the-unix-way"
    "https://superuser.com/questions/704870/mount-and-dismount-hard-drive-through-a-script-software"
    "https://www.uwe-sieber.de/drivetools_e.html"
    "https://www.download3k.com/articles/How-to-Remount-Safely-Removed-USB-Devices-without-Re-Plugging-Them-00258"
    "Microsoft Devcon: choco install devcon.portable   # Installs devon32 / devcon64   shorturl.at/psAL1"
    ""
    "Get-Partition -DiskNumber 1 | Set-Partition -NewDriveLetter Z"
    "https://www.windowscentral.com/how-assign-permanent-drive-letter-windows-10#assign_drive_letter_powershell"
    "https://winaero.com/blog/remove-drive-letter-windows-10/"
    "Remove-PartitionAccessPath -DiskNumber 1 -PartitionNumber 1 -Accesspath F:"
}


function Change-Scaling {
    # Scale monitor to 100% : https://www.mikaelgranberg.se/node/39
    cd 'HKCU:\Control Panel\Desktop'
    Set-ItemProperty -Path . -Name LogPixels -Value 96
    cmd.exe /c "shutdown /l"
    exit
}

# https://powershell.org/forums/topic/how-to-get-screen-resolution-with-dpi-scaling-in-a-remote-desktop-session/
function Get-Scaling {
    # https://hinchley.net/articles/get-the-scaling-rate-of-a-display-using-powershell/

    Add-Type @'
using System; 
using System.Runtime.InteropServices;
using System.Drawing;

public class DPI {  
  [DllImport("gdi32.dll")]
  static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

  public enum DeviceCap {
    VERTRES = 10,
    DESKTOPVERTRES = 117
  } 

  public static float scaling() {
    Graphics g = Graphics.FromHwnd(IntPtr.Zero);
    IntPtr desktop = g.GetHdc();
    int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
    int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);

    return (float)PhysicalScreenHeight / (float)LogicalScreenHeight;
  }
}
'@ -ReferencedAssemblies 'System.Drawing.dll'

    [Math]::round([DPI]::scaling(), 2) * 100
}



# http://www.bradleyschacht.com/collecting-server-performance-metrics-powershell/
# http://www.bradleyschacht.com/collecting-server-performance-metrics-performance-monitor/
# https://mcpmag.com/articles/2018/02/07/performance-counters-in-powershell.aspx
# https://www.red-gate.com/simple-talk/sysadmin/powershell/powershell-day-to-day-admin-tasks-monitoring-performance/
# https://www.datadoghq.com/blog/collect-windows-server-2012-metrics/
# Following from bradleyschact.com, saves counters to .csv for analysis
function Get-PerformanceCounters {
    cls
 
    $outputDirectory = "C:\0\Performance Counters" # Directory where the restult file will be stored.
    
    $computerName = ""    # Set the Computer from which to collect counters. Leave blank for local computer.
    $sampleInterval = 5   # Collection interval in seconds.
    $maxSamples = 240     # How many samples should be collected at the interval specified. Set to 0 for continuous collection.
    
    # Check to see if the output directory exists. If not, create it. 
    if (-not(Test-Path $outputDirectory))
        {
            Write-Host "Output directory does not exist. Directory will be created."
            $null = New-Item -Path $outputDirectory -ItemType "Directory"
            Write-Host "Output directory created."
        }
    
    # Strip the \ off the end of the directory if necessary. 
    if ($outputDirectory.EndsWith("\")) {$outputDirectory = $outputDirectory.Substring(0, $outputDirectory.Length - 1)}
    
    # Create the name of the output file in the format of "computer date time.csv".
    $outputFile = "$outputDirectory\$(if($computerName -eq ''){$env:COMPUTERNAME} else {$computerName}) $(Get-Date -Format "yyyy_MM_dd HH_mm_ss").csv"
    
    # Write the parameters to the screen.
    Write-Host "
    
    Collecting counters...
    Press Ctrl+C to exit."
    
    # Specify the list of performance counters to collect.
    $counters =
        @(`
        "\Processor(_Total)\% Processor Time" `
        ,"\Memory\Available MBytes" `
        ,"\Paging File(_Total)\% Usage" `
        ,"\LogicalDisk(*)\Avg. Disk Bytes/Read" `
        ,"\LogicalDisk(*)\Avg. Disk Bytes/Write" `
        ,"\LogicalDisk(*)\Avg. Disk sec/Read" `
        ,"\LogicalDisk(*)\Avg. Disk sec/Write" `
        ,"\LogicalDisk(*)\Disk Read Bytes/sec" `
        ,"\LogicalDisk(*)\Disk Write Bytes/sec" `
        ,"\LogicalDisk(*)\Disk Reads/sec" `
        ,"\LogicalDisk(*)\Disk Writes/sec"    
        )
    
    # Set the variables for the Get-Counter cmdlet.
    $variables = @{
        SampleInterval = $sampleInterval
        Counter = $counters
    }
    
    # Add the computer name if it was not blank.
    if ($computerName -ne "") {$variables.Add("ComputerName","$computerName")}

    # Either set the sample interval or specify to collect continuous.
    if ($maxSamples -eq 0) {$variables.Add("Continuous",1)}
    else {$variables.Add("MaxSamples","$maxSamples")}

    # Show the variables then execute the command while storing the results in a file.
    $variables
    Get-Counter @Variables | Export-Counter -FileFormat csv -Path $outputFile -Force
}


# Send email if disk space is low
# https://www.ryadel.com/en/runninglow-free-powershell-script-check-low-disk-space-send-email-alert/


function Create-Shortcut {
    # Template from Edwin: demonstrates how to create a .ico from a .exe, to use with shortcuts
    # $ExeFile $OutIcon should be parameters for function if use this more
    Add-Type -AssemblyName System.Drawing
    [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\calc.exe").ToBitmap().Save("C:\0\calc.bmp")  # extract icon to a .bmp
    $bmap = [System.Drawing.Bitmap]::FromFile("C:\0\calc.bmp")
    $icon = [System.Drawing.Icon]::FromHandle($bmap.GetHicon())                   # use it as an icon
    $icoFile = New-Object System.IO.FileStream("C:\0\calc.ico", 'OpenOrCreate')   # stream it to an .ico file
    $icon.Save($icoFile)
    $icoFile.Close()
    $icon.Dispose()
    $bmap.Dispose()
    
    # I use $makeIcon = $true and then extract the first icon with a try and catch because if
    # it has no icon then it should skip that part, you can work it out
    # It is also quite smart as if there are spaces it will handle itself.
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$pathyouwish\$shortcutNameYouwish.lnk")
    $Shortcut.TargetPath = $exeFile
    $Shortcut.WorkingDirectory = $workDir   # $workDir is always (Split-Path $exeFile)
    $Shortcut.Arguments = $argsyouwish
    # extra check, not needed most of the time, but I had that situation
    if ($makeIcon = $true) { $Shortcut.IconLocation = "$iconLocation\$iconName.ico" }
    $Shortcut.Save()
}


# https://gallery.technet.microsoft.com/scriptcenter/Fast-asynchronous-ping-IP-d0a5cf0e/view/Discussions/2
function Get-OSVersion {
	[CmdletBinding()] param( [Parameter(Position=0, Mandatory=$true)] [System.String] $IPAddress = '' )
    $Win32_OS = Get-WmiObject Win32_OperatingSystem -computer $IPAddress
    return $Win32_OS.Caption
}
function Set-Subnets {
    $subnets = @()
    $subnets += ,@("192.168.0.1","192.168.0.255")
    $subnets += ,@("192.168.1.1","192.168.1.255")
    return $subnets
}
function Start-SubnetQuery {
    $MyNetwork = Set-Subnets
    foreach ($subnet in $MyNetwork) {
        Write-Output "Pinging subnet: " $subnet[0]
        Query-Subnet -StartAddress $subnet[0] -EndAddress $subnet[1]
    }
}

function Format-TableCompact {
    [CmdletBinding()]  
    Param (
        [parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)] 
        [PsObject]$InputObject,
        [switch]$AppendNewline
    )

    # An alternative pipeline-aware table output that is more compact
    # If the data is sent through the pipeline, use the $input automatic variable
    # to collect it as an array:
    if ($PSCmdlet.MyInvocation.ExpectingInput) { $InputObject = @($Input) }
    # or use : $InputObject = $Input | ForEach-Object { $_ }

    $result = ($InputObject | Format-Table -AutoSize | Out-String).Trim()
    if($AppendNewline) { $result += [Environment]::NewLine }
    $result
}



# There is a module called WindowBox.RDP on PSGallery and it just contains this single function, so have reused it here
# One good thing is that it properly uses Invoke-CimMethod instead of WMI (WMI is redundant)
# Also could build a tool using these: https://discoposse.com/2012/10/20/finding-rdp-sessions-on-servers-using-powershell/
function Enable-RDP {
    # Enable RDP
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name fDenyTSConnections -Type DWord -Value 0

    # Disable Network Level Authentication
    $ts = Get-CimInstance -Namespace root\cimv2\terminalservices -ClassName Win32_TSGeneralSetting -Filter 'TerminalName = "RDP-Tcp"'
    $ts | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{UserAuthenticationRequired=0} | Out-Null

    # Enable RDP on the firewall
    Enable-NetFirewallRule -DisplayName 'Remote Desktop - User Mode (TCP-in)'
}

# Based on: Powershell Gallery Open-RDPGUI (2014)
# Should accept a hash table with IP : Username
# For a given IP-Username, have an encrypted password file in %temp%
# Leverage the 
function Open-RDPGUI {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

    # -------------------------
    #    Variable definition
    # ------------------------- 
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\mstsc.exe")
    $RDPServer = ""
    $Server = ""
    $LogFile = "RDP_GUI_$(Get-Date -Format yyyy-MM-dd_HH-mm-ss)"
    $LogFilePath = "C:\0\$LogFile.log"
    $Date = Get-Date
    $Font = New-Object System.Drawing.Font("Arial",8)
    $arrServers = @()
    $arrServers += "192.168.0.11"
    $arrServers += "192.168.0.26"
    $arrServers += "192.168.0.28"
    $arrServers += "192.168.0.29"
    
    # [void] $objListBox.Items.Add("192.168.0.11")
    # [void] $objListBox.Items.Add("192.168.0.26")
    # [void] $objListBox.Items.Add("192.168.0.28")
    # [void] $objListBox.Items.Add("192.168.0.29")
    
    # -------------------------
    #  RDP Connection Function
    # ------------------------- 
    Function Connect{
        mstsc.exe -v $RDPServer
        Out-File -FilePath $LogFilePath -InputObject "$Date - RDP connection was opened against $RDPServer" -Append
    }

    # -------------------------
    #      Creates Form
    # ------------------------- 
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Remote Connection"
    $objForm.Size = New-Object System.Drawing.Size(190,350)
    $objForm.StartPosition = "CenterScreen"
    # $objForm.Location.X = 10
    # $objForm.Location.Y = 22
    # $objForm.Name = "form1"
    # $objForm.DataBindings.DefaultDataSourceUpdateMode = 0
    # $objForm.add_Load($OnLoadForm_StateCorrection)
    # $objForm.ShowDialog()| Out-Null   # You can also see all of the available properties and methods by using:
    # $objForm | Get-Member
    # https://social.technet.microsoft.com/Forums/windowsserver/en-US/638fef57-f080-4ae5-ba60-fb947af8e453/position-a-windows-form-with-powershell?forum=winserverpowershell
    $objForm.Location = New-Object System.Drawing.Size(1170,200)
    $objForm.Opacity = 0.9
    $objForm.SizeGripStyle = "Hide"
    $objForm.BackColor = "CadetBlue"
    $objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $objForm.Icon = $Icon
    $objForm.Topmost = $true
    $objForm.KeyPreview = $True
    $Date = Get-Date
    $objForm.Add_KeyDown({
        if ($_.KeyCode -eq "Enter") {
            $Server=$objTextBox.Text
            mstsc /v $Server
            Out-File -filePath $LogFilePath -inputObject "$Date - RDP connection was opened against $Server" -append
        }
    })
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})

    # -------------------------
    #     Creates Label
    # -------------------------
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,15)
    $objLabel.Size = New-Object System.Drawing.Size(280,20) 
    $objLabel.Font = $Font
    $objLabel.Text = "Double click to connect:"

    # -------------------------
    #   Creates Second Label
    # -------------------------
    $objLabel1 = New-Object System.Windows.Forms.Label
    $objLabel1.Location = New-Object System.Drawing.Size(10,225)
    $objLabel1.Size = New-Object System.Drawing.Size(280,20)
    $objLabel1.Text = "Insert name/IP and press Enter:"

    # -------------------------
    #    Creates List Box
    # -------------------------
    $objListBox = New-Object System.Windows.Forms.ListBox
    $objListBox.Location = New-Object System.Drawing.Size(10,40)
    $objListBox.Size = New-Object System.Drawing.Size(160,20)
    $objListBox.Height = 180
    $objListBox.Add_DoubleClick({$RDPServer = $objlistbox.SelectedItem;Connect})

    # -------------------------
    #   Add Items to List Box
    # -------------------------
    foreach ($i in $arrServers){
        [void] $objListBox.Items.Add($i)
    }

    # -------------------------
    #     Creates Text Box
    # -------------------------
    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.Location = New-Object System.Drawing.Size(10,250) 
    $objTextBox.Size = New-Object System.Drawing.Size(160,20) 

    # -------------------------
    #  Creates Open Log Button
    # ------------------------- 
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(10,285)
    $OKButton.Size = New-Object System.Drawing.Size(160,23)
    $OKButton.Text = "Open Log File"
    $OKButton.Add_Click({Invoke-Item $LogFilePath})

    # -------------------------------
    #  Add objects and Activate Form
    # ------------------------------- 
    $objForm.Controls.Add($objLabel)
    $objForm.Controls.Add($objListBox)
    $objForm.Controls.Add($objTextBox)
    $objForm.Controls.Add($objLabel)
    $objForm.Controls.Add($objLabel1)
    $objForm.Controls.Add($OKButton)
    $objForm.Add_Shown({$objForm.Activate()})
    $objForm.ShowDialog()
}

# username and password are optional, if they are provided, save an encrypted password to $env:TEMP  ps_<hostname>_<username>_password.txt
function rdphalf ($hostname, $username) {
    <#
    .SYNPPSIS
    RDP (Remote Desktop Protcol) Automatic Logon and properly resize to exactly half-size of screen for split screen work
    ToDo: Maximise window *or* snap to left after opening.

    rpdhalf <host> <username>
    Password will be prompted for if required and saved in a secure credential file
    #>

    $passcred = "$($env:TEMP)\ps_rdp_$($hostname)_$($username).password"
    if (!(Test-Path $passcred)) {
        Read-Host "Enter the password for $username to RDP onto $hostname" -AsSecureString | ConvertFrom-SecureString | Out-File $passcred
        Write-Host "Password is encrypted to: $passcred"
        Write-Host "These credentials will be used for all future access to $hostname with this tool."
        Write-Host "Delete the secure password file and rerun this tool to regenerate a new password."
    }

    # Assigning $cred here means that running mstsc will automatically use the last credentials (i.e. won't ask for password)
    $password = Get-Content $passcred | ConvertTo-SecureString
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    # https://adamtheautomator.com/powershell-get-credential/
    # https://stackoverflow.com/questions/6239647/using-powershell-credentials-without-being-prompted-for-a-password
    # Start-Process -WindowStyle Hidden "C:\Program Files\Internet Explorer\iexplore.exe" "www.google.com"   #
    
    # Get exactly half-width and full height (minus the Start Bar height)
    $width = (Get-WmiObject -Class Win32_DesktopMonitor | Select-Object ScreenWidth).ScreenWidth
    $height = (Get-WmiObject -Class Win32_DesktopMonitor | Select-Object ScreenHeight).ScreenHeight
    $width = ($width / 2) - 5
    $height = $height - 107   # default Start Bar is 72 or 74 high and short Start Bar is 24 high, so 900 - 72 = 828 for 900 height
    # 107 is ad hoc, need to get this calculation better
    Start-Process "mstsc" "/v:$hostname /w:$width /h:$height"
}

# c:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -nolog -command cmdkey /generic:TERMSRC/some_unc_path /user:username /pass:pa$$word; mstsc /v:some_unc_path
# Remote Desktop Manager allows single click access to multiple servers (try this)

# write-output "Connecting to Computername"
# $Server   = "username"
# $User     = "computername\username"
# $Password = "password"
# 
# cmdkey /delete:"$Server" # probably not needed, just clears the credentials
# cmdkey /generic:"$Server" /user:"$user" /pass:"$password"
# 
# mstsc /v:"$Server" /admin # /admin probably not needed either

# "Identity is not fully verified. Please enter new credentials" After google search could be a permissions issue
# with saved credentials passing through RDP. From what I have found I will need to edit the group policy underer:
#   Computer Configuration -> Administrative Templates -> System -> Credentials Delegation 

function Fix-WorkLaptop {
    # Remove some bloat put on by company, kill some processes, reset some things
    # Some things might require admin to kill (e.g. Tight VNC Server)
    # Note: Not a generic function, has things customised for my own laptop
    kill -name workpace -force -EA Silent    # WorkPace. Aannoying tool to make you stretch your back etc
    kill -name Docker* -force -EA Silent     # Docker Desktop & Docker.Watchguard
    kill -name Steam* -force -EA Silent      # Steam
    kill -name AutoHotkey -force -EA Silent  # Restart this with Main and Main-ING to fix any issues
    kill -name Rainmeter -force -EA Silent   # Rainmeter, until I know how to configure properly
    kill -name "My Epson*" -force -EA Silent # My Epson printer tools
    
    $user = [Security.Principal.WindowsIdentity]::GetCurrent()

    function DoIfAdmin ($job) {
        if ((New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) { $job }
        else { "`nCould not run '$job' as this is a non-elevated account.`n" }
    }

    DoIfAdmin "kill -name tvnserver -force -EA Silent"   # Tight VNC Server
    DoIfAdmin "Disable-BitLockerEncryption | Out-Null"   # Allow all USB devices read/write, Note that encyption will re-enable in time when the policy is reapplied
        # Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker -Name PreventDeviceEncryption -Value 1 -Force -EA Silent
        # Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Policies\Microsoft\FVE -Name RDVDenyWriteAccess -Value 0 -Force -EA Silent
    
    # Apply some Startup Tasks, but test if available first
    # C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden -file "C:\0\SysTray\SysTray.ps1"
   
    # Start Main.ahk
    if (Test-Path C:\0\Scripts_AHK\Basics.ahk) { Start-Process C:\0\Scripts_AHK\Basics.ahk }
    if (Test-Path C:\0\Scripts_AHK\Main-ING.ahk) { Start-Process C:\0\Scripts_AHK\Main-ING.ahk }
    if (Test-Path "C:\0\Scripts_AHK\WOTR Toolkit.ahk") { Start-Process "C:\0\Scripts_AHK\WOTR Toolkit.ahk" }
    
    # Option to gracefully shutdown process   https://stackoverflow.com/questions/28481811/how-to-correctly-check-if-a-process-is-running-and-stop-it
    # $firefox = Get-Process firefox -ErrorAction SilentlyContinue
    # if ($firefox) {
    #       $firefox.CloseMainWindow()   # try gracefully first
    #       Sleep 5   # kill after five seconds
    #       if (!$firefox.HasExited) { $firefox | Stop-Process -Force }
    # }
    # Remove-Variable firefox
}

Set-Alias Fix-ING Fix-WorkLaptop

function Restart-AutoHotkey {
    # Look for my main projects and start or restart them as required
    kill -name AutoHotkey -force -EA Silent
    if (Test-Path "C:\0\Scripts_AHK\Basics.ahk") { Start-Process "C:\0\Scripts_AHK\Basics.ahk" }
    if (Test-Path "C:\0\Scripts_AHK\Main-ING.ahk") { Start-Process "C:\0\Scripts_AHK\Main-ING.ahk" }
    if (Test-Path "C:\0\Scripts_AHK\WOTR Toolkit.ahk") { Start-Process "C:\0\Scripts_AHK\WOTR Toolkit.ahk" }
    if (Test-Path "D:\0 Cloud\OneDrive\0_Scripts_AutoHotkey\Basics-AHK\Basics.ahk") { Start-Process "D:\0 Cloud\OneDrive\0_Scripts_AutoHotkey\Basics-AHK\Basics.ahk" }
}

function Install-ModuleFromPSGallery {
    [CmdletBinding()] [OutputType('System.Management.Automation.PSModuleInfo')]
    param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]                                    $Name,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidateScript({ Test-Path $_ })] $Destination
    )

    # Force installtaion to a specific location (usually the users own Modules path)

    if (!(Test-Path $UserModulesPath)) { New-Item $UserModulesPath -ItemType Directory -Force }

    # If the Module is installed at a network location path, remove it and move to user Modules path
    # Nothing will happen here unless working on work laptop with user shares on UNC paths.
    if (($Profile -like "\\*") -and (Test-Path (Join-Path $UserModulesPath $Name))) {
        if (Test-Administrator -eq $true) {
            "remove module from network share and move to $Destination"
            Uninstall-Module $Name -Force -Verbose
            # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share if in use so use Save-Module
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
        }
        else {
            Write-Host "Module found on network share module path but need to be administrator and connected to VPN" -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "to correctly move Modules into the users module folder on C:\" -ForegroundColor Yellow -BackgroundColor Black
        }
    }
    elseif (Test-Path (Join-Path $AdminModulesPath $Name)) {
        if (Test-Administrator -eq $true) {
            "remove module from $AdminModulesPath and move to $Destination"
            Uninstall-Module $Name -Force -Verbose
            # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share if in use so use Save-Module
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
        }
        else {
            Write-Host "Module found on in Admin Modules folder: $(split-path $AdminModulesPath) C:\Program Files\WindowsPowerShell\Modules." -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "Need to be Admin to correctly move Modules into the users module folder on C:\" -ForegroundColor Yellow -BackgroundColor Black
        }
    }
    # Get-InstalledModule   # Shows only the Modules installed by PowerShellGet.
    # Get-Module            # Gets the modules that have been imported or that can be imported into the current session.
    elseif (Test-Path (Join-Path $Destination $Name)) {
        # https://stackoverflow.com/questions/48424152/compare-system-version-in-powershell
        # To use the repository, you either need PowerShell 5 or install the PowerShellGet module manually (which is
        # available for download on powershellgallery.com) to get Find/Save/Install/Update/Remove-Script for Modules.
        # https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/getting-latest-powershell-gallery-module-version
        # https://stackoverflow.com/questions/52633919/powershell-sort-version-objects-descending
        # "1.." -match "\b\d(\.\d{0,5}){0,3}\d$"
        # https://techibee.com/powershell/check-if-a-string-contains-numbers-in-it-using-powershell/2842
        $ModVerLocal = (Get-Module $Name -ListAvailable -EA Silent).Version
        $ModVerOnline = Get-PublishedModuleVersion $Name
        $ModVerLocal = "$(($ModVerLocal).Major).$(($ModVerLocal).Minor).$(($ModVerLocal).Build)"      # reuse the [version] variable as a [string]
        $ModVerOnline = "$(($ModverOnline).Major).$(($ModverOnline).Minor).$(($ModverOnline).Build)"  # reuse the [version] variable as a [string]
        # if ($ModuleVersionOnline -ne "") { $ModuleVersionOnline = "$($ModuleVersionOnline.split(".")[0]).$($ModuleVersionOnline.split(".")[1]).$($ModuleVersionOnline.split(".")[2])" }
        echo "Local Version:  $ModVerLocal"
        echo "Online Version: $ModVerOnline"
        if ($ModVerLocal -eq $ModVerOnline) {
            echo "$Name is installed and latest version, nothing to do!"
        }
        else {
            if ([bool](Get-Module $Name) -eq $true) { Uninstall-Module $Name -Force -Verbose }
            rm (Join-Path $Destination $Name) -Force -Recurse -Verbose
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination -Force -Verbose   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name) -Force -Verbose
        }
    }
    else {   # Final case is no module is in network share, or local admin modules, or local user modules so now just install it
        Get-PublishedModuleVersion $Name
        Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination -Force -Verbose   # Install the module to the custom destination.
        Import-Module -FullyQualifiedName (Join-Path $Destination $Name) -Force -Verbose
    }

    # Finally, output the Path to the newly installed module and the functions contained in it
    (Get-Module $Name | select Path).Path
    $out = ""; foreach ($i in (Get-Command -Module $Name).Name) {$out = "$out, $i"} ; "" ; Write-Wrap $out.trimstart(", ") ; ""
    # return (Get-Module)
}

function Pull-Gist {

    $jumpfrom = Get-Location   # Save the current location
    if (Test-Path "$HomeFix\Gist") { Set-Location "$HomeFix\Gist" }
    if (Test-Path "C:\0\Gist") { Set-Location "C:\0\Gist" }
    if (Test-Path "D:\Gist") { Set-Location "D:\Gist" }
    if (Test-Path "D:\0 Cloud\OneDrive\Gist") { Set-Location "D:\0 Cloud\OneDrive\Gist" }

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls } catch { }   # Windows 7 compatible
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }   # Windows 10 compatible
    Clear-DnsClientCache
    [System.Net.ServicePointManager]::DnsRefreshTimeout = 0;

    Write-Host ""
    Write-Host "Note: none of the below will work if connected to Endopint."
    Write-Host "or to the ENTER network (reconnect to GUEST)."
    Write-Host ""
    pause

    $now = Get-Date -format "yyyy-MM-dd__HH-mm-ss"
    # $pt = ".\Pull-Temp"
    # if (! (Test-Path $pt)) { md $pt }

    if (Test-Path .\BeginSystemConfig.ps1) { mv .\BeginSystemConfig.ps1 .\BeginSystemConfig_$($now).ps1 }
    if (Test-Path .\ProfileExtensions.ps1) { mv .\ProfileExtensions.ps1 .\ProfileExtensions_$($now).ps1 }
    if (Test-Path .\Custom-Tools.psm1)     { mv .\Custom-Tools.psm1 .\Custom-Tools_$($now).psm1 }
    if (Test-Path .\Gist-Push.ps1)         { mv .\Gist-Push.ps1 .\Gist-Push_$($now).ps1 }

    iwr 'https://gist.github.com/roysubs/61ef677591f22927afadc9ef2b657cd9/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\BeginSystemConfig.ps1
    iwr 'https://gist.github.com/roysubs/c37470c98c56214f09f0740fcb21ec4f/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\ProfileExtensions.ps1
    iwr 'https://gist.github.com/roysubs/5c6a16ea0964cf6d8c1f9eed7103aec8/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\Custom-Tools.psm1
    iwr 'https://gist.github.com/roysubs/908525ae135e7d31a4fd13bd111b50e9/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\Push-Gist.ps1

    if (Test-Path $jumpfrom) { Set-Location $jumpfrom }
}

function Push-Gist {

    $jumpfrom = Get-Location   # Save the current location
    if (Test-Path "$HomeFix\Gist") { Set-Location "$HomeFix\Gist" }
    if (Test-Path "C:\0\Gist") { Set-Location "C:\0\Gist" }
    if (Test-Path "D:\Gist") { Set-Location "D:\Gist" }
    if (Test-Path "D:\0 Cloud\OneDrive\Gist") { Set-Location "D:\0 Cloud\OneDrive\Gist" }

    # Need to make sure using TLS for connection to GitHub
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls } catch { }   # Windows 7 compatible
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }   # Windows 10 compatible

    # The endpoint is cached when connecting to GitHub, so your updates may be visible with a delay (10 sec to 2 min in my experience).
    # https://stackoverflow.com/questions/46073096/is-there-a-permalink-to-the-latest-version-of-gist-files
    # How long it will take for the cached version to be updated with the newest changes?
    # It looks like you can cache-bust by attaching a query string to the url @ nietaki Apr 18 '18 at 21:49  (I never got this working)
    # e.g. https://gist.githubusercontent.com/mwek/9962f97f3bde157fd5dbd2b5dd0ec3ca/raw/user.js?cachebust=dkjflskjfldkf

    if (!(Test-Path '.\Push-Gist-Secure-Password.txt')) {
        Read-Host "Enter the github password for roysubs" -AsSecureString | ConvertFrom-SecureString | Out-File .\Push-Gist-Secure-Password.txt
        Write-Host "Password now saved securely to .\Push-Gist-Secure-Password.txt.`nIf you want to regenerate a new password, delete that file."
    }
    $username = 'roysubs'
    $password = Get-Content '.\Push-Gist-Secure-Password.txt' | ConvertTo-SecureString
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    # https://adamtheautomator.com/powershell-get-credential/
    # https://stackoverflow.com/questions/6239647/using-powershell-credentials-without-being-prompted-for-a-password

    try {
        iwr 'https://gist.github.com/roysubs/61ef677591f22927afadc9ef2b657cd9/raw'
    }
    catch {
        Write-Host "Invoke-WebRequest failed. You might be behind a firewall / VPN or`nthe Internet Explorer engine might not be fully initialised.`nPlease correct this then retry."
        Start-Process -WindowStyle Hidden "C:\Program Files\Internet Explorer\iexplore.exe" "www.google.com"
        # $IE=new-object -com internetexplorer.application
        # $IE.navigate2("www.microsoft.com")
        # $IE.visible=$true
        # For your reference:
        # Controlling Internet Explorer object from PowerShell
        # http://blogs.msdn.com/powershell/archive/2006/09/10/controlling-internet-explorer-object-from-powershell.aspx
        # https://social.technet.microsoft.com/Forums/ie/en-US/e54555bd-00bb-4ef9-9cb0-177644ba19e2/how-to-open-url-through-powershell
        try {
            iwr 'https://gist.github.com/roysubs/61ef677591f22927afadc9ef2b657cd9/raw'
        }
        catch {
            throw "Invoke-WebRequest failed again. Makes sure that Internet Explorer is fully initialised and check internet / VPN."
            Start-Process -WindowStyle Hidden "C:\Program Files\Internet Explorer\iexplore.exe" "www.google.com"
        }
    }

    # Test if Posh-Gist is installed
    if (!(Test-Path("C:\Program Files\WindowsPowerShell\Modules\posh-gist\*\posh-gist.psm1"))) { Install-Module Posh-Gist }

    # Securely upload the Gists, testing existence and non-zero size
    if (Test-Path(".\BeginSystemConfig.ps1")) {
        if ((Get-Item ".\BeginSystemConfig.ps1").length -gt 0kb) {
            Update-Gist -Credential $cred -Id 61ef677591f22927afadc9ef2b657cd9 -Update .\BeginSystemConfig.ps1
        }
    }

    if (Test-Path(".\ProfileExtensions.ps1")) {
        if ((Get-Item ".\ProfileExtensions.ps1").length -gt 0kb) {
            Update-Gist -Credential $cred -Id c37470c98c56214f09f0740fcb21ec4f -Update .\ProfileExtensions.ps1
        }
    }

    if (Test-Path(".\Custom-Tools.psm1")) {
        if ((Get-Item ".\Custom-Tools.psm1").length -gt 0kb) {
            Update-Gist -Credential $cred -Id 5c6a16ea0964cf6d8c1f9eed7103aec8 -Update .\Custom-Tools.psm1
        }
    }

    if (Test-Path(".\Gist-Push.ps1")) {
        if ((Get-Item ".\Gist-Push.ps1").length -gt 0kb) {
            Update-Gist -Credential $cred -Id 908525ae135e7d31a4fd13bd111b50e9 -Update .\Gist-Push.ps1
        }
    }

    # Try clearing the DNS client cache to resolve the endpoint caching (this does not work in my experience)
    Clear-DnsClientCache
    [System.Net.ServicePointManager]::DnsRefreshTimeout = 0;

    if (!(Test-Path(".\temp"))) { New-Item -Type Directory ".\temp" }
    iwr 'https://gist.github.com/roysubs/61ef677591f22927afadc9ef2b657cd9/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\temp\BeginSystemConfig.ps1
    iwr 'https://gist.github.com/roysubs/c37470c98c56214f09f0740fcb21ec4f/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\temp\ProfileExtensions.ps1
    iwr 'https://gist.github.com/roysubs/5c6a16ea0964cf6d8c1f9eed7103aec8/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\temp\Custom-Tools.psm1
    iwr 'https://gist.github.com/roysubs/908525ae135e7d31a4fd13bd111b50e9/raw' -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File .\temp\Gist-Push.ps1

    if (Test-Path $jumpfrom) { Set-Location $jumpfrom }
}


function Get-DiskSpeed {

    param( [string]$driveLetter = "C:", [int]$sampleSize = 5 )
    
    # Script is from: https://kimconnect.com/powershell-benchmark-disk-speed/
    # Requires diskspd.exe : https://gallery.technet.microsoft.com/DiskSpd-A-Robust-Storage-6ef84e62
    # https://www.trishtech.com/2017/01/benchmark-storage-disks-with-microsoft-diskspeed-tool/

    # Example of manual pull and extract for diskspd.exe
    # $url = "https://gallery.technet.microsoft.com/DiskSpd-A-Robust-Storage-6ef84e62/file/199535/2/DiskSpd-2.0.21a.zip"
    # $FileName = ($url -split "/")[-1]   # Could also use:  $url -split "/" | select -last 1   # 'hi there, how are you' -split '\s+' | select -last 1
    # (New-Object System.Net.WebClient).DownloadString($url) | Out-File $((Get-Lotcation).Path)\$FileName
    # Below script gets the tool using Chocolatey though which is easier.

    if($driveLetter.Length -eq 1) { $driveLetter+=":"; }
    Write-Host "Obtaining disk speed of $driveLetter ..."
    
    function isPathWritable {
        param($testPath)
        # Create random test file name
        $tempFolder = $testPath + "\getDiskSpeed\"
        $filename = "diskSpeedTest-"+[guid]::NewGuid()
        $tempFilename = (Join-Path $tempFolder $filename)
        New-Item -ItemType Directory -Path $tempFolder -Force -EA SilentlyContinue | Out-Null
        try { 
            # Try to add a new file
            # New-Item -ItemType Directory -Path $tempFolder -Force -EA SilentlyContinue
            [io.file]::OpenWrite($tempFilename).Close()
            # Write-Host -ForegroundColor Green "$testPath is writable."         
            # Delete test file after done
            # Remove-Item $tempFilename -Force -ErrorAction SilentlyContinue 
                
            # Set return value
            $feasible=$true;
        }
        catch {
            # Return 'false' if there are errors
            $feasible=$false;
        }
        return $feasible;
    }

    # Check if input is a valid drive letter
    function validatePath {
        param( [string]$path = $driveLetter )

        if (Test-Path $path -EA SilentlyContinue) {
            $regexValidDriveLetters = "^[A-Za-z]\:{0,1}$"
            $validLocalPath=$path.SubString(0,2) -match $regexValidDriveLetters

            if ($validLocalPath) {
                $GLOBAL:localPath=$true
                Write-Host "Validating path... Local directory detected."
                
                $volumeName = if($driveLetter.Length -le 2) { $driveLetter+"\" } else { $driveLetter.Substring(0,3) }
                $GLOBAL:clusterSize = (Get-WmiObject -Class Win32_Volume | Where-Object { $_.Name -eq $volumeName }).BlockSize
                Write-Host "Cluster size detected as $clusterSize."

                $driveLettersOnThisComputer = ls function:[A-Z]: -n | ? { test-path $_ }
                if (!($driveLettersOnThisComputer -contains $path.SubString(0,2))) {
                    Write-Host "The provided local path's first 2 characters do not match any volumes in this system."
                    return $false
                }
                return $(isPathWritable $path)
            } else {
                $regexUncPath="^\\(?:\\[^<>:`"/\\|?*]+)+$"
                if ($path -match $regexUncPath) {
                    $GLOBAL:localPath=$False;
                    write-Host "UNC directory detected."
                    return $(isPathWritable $path)
                } else {
                    Write-Host "The provided path does not match a UNC pattern nor a local drive."
                    return $false
                }
            }
        } else {
            Write-Host "The path $path currently does NOT exist"
            Return $false
        }
    }
    
    if (validatePath) {
        # Set variables
        $tempDirectory = "$driveLetter`\getDiskSpeed"
        # New-Item -ItemType Directory -Force -Path $tempDirectory|Out-Null
        $testFile = "$tempDirectory`\testfile.dat"
        $processors = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors
           
        # Use Chocolatey to ensure that diskspd.exe is available in the system
        $diskSpeedUtilityAvailable = Get-Command diskspd.exe -EA SilentlyContinue
        if (!($diskSpeedUtilityAvailable)) {
            if (!(Get-Command choco.exe -ErrorAction SilentlyContinue)) {
                Set-ExecutionPolicy Bypass -Scope Process -Force
                iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
            }
            choco install diskspd -y --ignore-checksums
            refreshenv
        }
    
        function getIops {
        # Sometimes, the test result throws this error "diskspd Error opening file:" if no switches were used
        # The work around is to specify more parameters
        # Other variations:
        # $testResult=diskspd.exe-d1 -o4 -t4 -b8k -r -L -w50 -c1G $testFile
        # $testResult=diskspd.exe -b4K -t1 -r -w50 -o32 -d10 -c8192 $testFile
        # Note: remove the -c option to avoid this error when running with unprivileged accounts
        # diskspd.exe : WARNING: Error adjusting token privileges for SeManageVolumePrivilege (error code: 1300)
        
        try {
            if ($localPath) {
                #$expression="diskspd.exe -b8k -d1 -o$processors -t$processors -r -L -w25 -c1G $testfile"
                $expression = "Diskspd.exe -b$clusterSize -d1 -h -L -o$processors -t1 -r -w30 -c1G $testfile  2>&1"
            } else {
                $expression = "Diskspd.exe -b8K -d1 -h -L -o$processors -t1 -r -w30 -c1G $testfile 2>&1"
            }
            
            # Write-Host $expression
            $testResult = Invoke-Expression $expression
            <#
            diskspd.exe -b8k -d1 -o4 -t4 -r -L -w25 -c1G $testfile
            8K block size; 1 second random I/O test;4 threads; 4 outstanding I/O operations;
            25% write (implicitly makes read 75% ratio); 
            #>
        }
        catch {
            $errorMessage = $_.Exception.Message
            $failedItem = $_.Exception.ItemName
            Write-Host "$errorMessage $failedItem";
            continue;
        }
        $x = $testResult | select-string -Pattern "total*" -CaseSensitive | select-object -First 1 | Out-String
        $iops = $x.split("|")[-3].Trim()
        #$mebibytesPerSecond=$x.split("|")[-4].Trim()            
        return $iops
    }
    
    function selectHighIops {
        $testArray = @()
        for($i = 1; $i -le $sampleSize; $i++) {
            try {
                $iops=getIops
                write-host "$i of $sampleSize`: $iops IOPS"
                $testArray+=$iops
            }
            catch {
                $errorMessage = $_.Exception.Message
                $failedItem = $_.Exception.ItemName
                Write-Host "$errorMessage $failedItem"
                break
            }
        }
        $highestResult = ($testArray | measure -Maximum).Maximum
        return $highestResult
    } 
    
    # Trigger several tests and select the highest value
    $selectedIops = selectHighIops
    # Cleanup
    # cmd /c rd $tempDirectory

    function isFileLocked {
        param ( $file = $(New-Object System.IO.FileInfo $testFile) )
        if (Test-Path $testFile) {
            try {
                $fileHandle = $file.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                if ($fileHandle) {
                    # File handle is open, which means file is not locked
                    $fileHandle.Close()
                }
                return $false
            }
            catch {
                # file is locked
                return $true
            }
        } else {
            return $false
        }
    }
    
    do {
        sleep 1
        isFileLocked | out-null
    } until (!(isFileLocked))
    Remove-Item -Recurse -Force $tempDirectory
    
    $mebibytesPerSecond = [math]::round($(([int]$selectedIops) / 128), 2)
    return "Highest: $selectedIops IOPS ($mebibytesPerSecond MiB/s)"
    } else {
        return "Cannot get disk speed"
    }
}

<#
(gwmi -Class win32_volume -Filter "DriveType!=5" -ea stop| ?{$_.DriveLetter -ne $isnull}|`
    Select-object @{Name="Letter";Expression={$_.DriveLetter}},`
    @{Name="Label";Expression={$_.Label}},`
    @{Name="Capacity";Expression={"{0:N2} GiB" -f ($_.Capacity/1073741824)}},`
    @{Name = "Available"; Expression = {"{0:N2} GiB" -f ($_.FreeSpace/1073741824)}},`
    @{Name = "Utilization"; Expression = {"{0:N2} %" -f  ((($_.Capacity-$_.FreeSpace) / $_.Capacity)*100)}},`
    @{Name = "diskBrand"; Expression = {getDiskSpeed $_.DriveLetter}},`
    @{Name = "diskSpeed"; Expression = {getDiskSpeed $_.DriveLetter}}`
    | ft -autosize | Out-String).Trim()
#>

function changeAutoLock {
    # Disable_AutoLock.ps1
    param ( [Switch]$disable, [Switch]$enable )

    $registryHive = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
    $registryKey = "NoLockScreen"
    $keyItem = Get-ItemProperty -Path $registryHive -Name $registryKey -ErrorAction SilentlyContinue 
    $keyValue  = $keyItem.NoLockScreen 
 
	if ($enable)
	{
		if ($keyValue) {
			#Remove item property 
			Remove-Item -Path $registryHive -Recurse -Confirm:$false  | Out-Null 
			Write-Host "Enabled lock screen successfully."
		} else {
			Write-Host "Lock screen has already been enabled prior."
		}
		
	} else {
	    if ($keyValue) {
		    Write-Host  "Lock screen has already been disabled."
	    } else {
		    New-Item -Path $registryHive -ErrorAction SilentlyContinue  | Out-Null 
		    New-ItemProperty -Path $registryHive -Type "DWORD" -Name "NoLockScreen"  -Value 1 | Out-Null 
		    Write-Host "Disabled lock screen successfully."
	    }
	}
}
# changeAutoLock -disable

# benchmark.psm1
# Exports: Benchmark-Command
function Benchmark-Command ([ScriptBlock]$Expression, [int]$Samples = 1, [Switch]$Silent, [Switch]$Long) {
    <#
    .SYNOPSIS
    Runs the given script block and returns the execution duration. http://zduck.com/2013/benchmarking-with-Powershell/
    Hat tip to StackOverflow. http://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
    Benchmark-Command { ping -n 1 google.com } -Samples 50 -Silent
      
    .EXAMPLE
    Benchmark-Command { ping -n 1 google.com } -Samples 50 -Silent
    #>
    
    $timings = @()
    do {
        $sw = New-Object Diagnostics.Stopwatch
        if ($Silent) {
            $sw.Start()
            $null = & $Expression
            $sw.Stop()
            Write-Host "." -NoNewLine
        }
        else {
            $sw.Start()
            & $Expression
            $sw.Stop()
        }
        $timings += $sw.Elapsed
        $Samples--
    }
    while ($Samples -gt 0)
    Write-Host
    
    $stats = $timings | Measure-Object -Average -Minimum -Maximum -Property Ticks
    
    # Print the full timespan if the $Long switch was given.
    if ($Long) {  
        Write-Host "Avg: $((New-Object System.TimeSpan $stats.Average).ToString())"
        Write-Host "Min: $((New-Object System.TimeSpan $stats.Minimum).ToString())"
        Write-Host "Max: $((New-Object System.TimeSpan $stats.Maximum).ToString())"
    }
    else {
        # Otherwise just print the milliseconds which is easier to read.
        Write-Host "Avg: $((New-Object System.TimeSpan $stats.Average).TotalMilliseconds)ms"
        Write-Host "Min: $((New-Object System.TimeSpan $stats.Minimum).TotalMilliseconds)ms"
        Write-Host "Max: $((New-Object System.TimeSpan $stats.Maximum).TotalMilliseconds)ms"
    }
}
    # Export-ModuleMember Benchmark-Command
    # Could run with this script to get performance over time
    # https://www.powershellbros.com/run-script-to-check-cpu-and-memory-utilization/


    # https://gallery.technet.microsoft.com/scriptcenter/Powershell-script-to-test-5d417634
    # 8/19/2014:This script has been retired. Its functionality has been included in the Test-SBDisk function, part of the SBTools module.This script tests disk IO performance by creating random files on the target disk and measuring IO performance  Warning:This script will delete

    # https://github.com/nullzeroio/PowerShell/blob/master/Invoke-CPUStressTest.ps1
    # https://superuser.com/questions/396501/how-can-i-produce-high-cpu-load-on-windows

    # winsat prepop
    # Get-WmiObject -class Win32_WinSAT







# If you want to run this on PowerShell < 3.0 use
# New-Object -TypeName PSObject -Property @{ ... } wherever it says [PSCustomObject]@{ ... }
# and change the -version value for 'requires' to 2     #""" requires -version 3 """"
# https://stackoverflow.com/questions/54041911/fast-registry-searcher-in-powershell
function Search-Registry {
    <#
        .SYNOPSIS
            Searches the registry on one or more computers for a specified text pattern.
        .DESCRIPTION
            Searches the registry on one https://www.youtube.com/watch?v=H3FP1eim4hYor more computers for a specified text pattern. 
            Supports searching for any combination of key names, value names, and/or value data. 
            The search pattern is either a regular expression or a wildcard pattern using the 'like' operator.
            (both are case-insensitive)
        .PARAMETER ComputerName
            (Required) Searches the registry on the specified computer(s). This parameter supports piped input.
        .PARAMETER Pattern
            (Optional) Searches using a wildcard pattern and the -like operator.
            Mutually exclusive with parameter 'RegexPattern'
        .PARAMETER RegexPattern
            (Optional) Searches using a regular expression pattern.
            Mutually exclusive with parameter 'Pattern'
        .PARAMETER Hive
            (Optional) The registry hive rootname.
            Can be any of 'HKEY_CLASSES_ROOT','HKEY_CURRENT_CONFIG','HKEY_CURRENT_USER','HKEY_DYN_DATA','HKEY_LOCAL_MACHINE',
                          'HKEY_PERFORMANCE_DATA','HKEY_USERS','HKCR','HKCC','HKCU','HKDD','HKLM','HKPD','HKU'
            If not specified, the hive must be part of the 'KeyPath' parameter.
        .PARAMETER KeyPath
            (Optional) Starts the search at the specified registry key. The key name contains only the subkey.
            This parameter can be prefixed with the hive name. 
            In that case, parameter 'Hive' is ignored as it is then taken from the given path.
            Examples: 
              HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall
              HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall
              Software\Microsoft\Windows\CurrentVersion\Uninstall
        .PARAMETER MaximumResults
            (Optional) Specifies the maximum number of results per computer searched. 
            A value <= 0 means will return the maximum number of possible matches (2147483647).
        .PARAMETER SearchKeyName
            (Optional) Searches for registry key names. You must specify at least one of -SearchKeyName, -SearchPropertyName, or -SearchPropertyValue.
        .PARAMETER SearchPropertyName
            (Optional) Searches for registry value names. You must specify at least one of -SearchKeyName, -SearchPropertyName, or -SearchPropertyValue.
        .PARAMETER SearchPropertyValue
            (Optional) Searches for registry value data. You must specify at least one of -SearchKeyName, -SearchPropertyName, or -SearchPropertyValue.
        .PARAMETER Recurse
            (Optional) If set, the function will recurse the search through all subkeys found.
        .OUTPUTS
            PSCustomObjects with the following properties:

              ComputerName     The computer name where the search was executed
              Hive             The hive name used in Win32 format ("CurrentUser", "LocalMachine" etc)
              HiveName         The hive name used ("HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE" etc.)
              HiveShortName    The abbreviated hive name used ("HKCU", "HKLM" etc.)
              Path             The full registry path ("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall")
              SubKey           The subkey without the hive ("Software\Microsoft\Windows\CurrentVersion\Uninstall")
              ItemType         Informational: describes the type 'RegistryKey' or 'RegistryProperty'
              DataType         The .REG formatted datatype ("REG_SZ", "REG_EXPAND_SZ", "REG_DWORD" etc.). $null for ItemType 'RegistryKey'
              ValueKind        The Win32 datatype ("String", "ExpandString", "DWord" etc.). $null for ItemType 'RegistryKey'
              PropertyName     The name of the property. $null for ItemType 'RegistryKey'
              PropertyValue    The value of the registry property. $null for ItemType 'RegistryKey'
              PropertyValueRaw The raw, unexpanded value of the registry property. $null for ItemType 'RegistryKey'

              The difference between 'PropertyValue' and 'PropertyValueRaw' is that in 'PropertyValue' Environment names are expanded
              ('%SystemRoot%' in the data gets expanded to 'C:\Windows'), whereas in 'PropertyValueRaw' the data is returned as-is.
              (Environment names return as '%SystemRoot%')

        .EXAMPLE
            Search-Registry -Hive HKLM -KeyPath SOFTWARE -Pattern $env:USERNAME -SearchPropertyValue -Recurse -Verbose

            Searches HKEY_LOCAL_MACHINE on the local computer for registry values whose data contains the current user's name.
            Searches like this can take a long time and you may see warning messages on registry keys you are not allowed to enter.
        .EXAMPLE
            Search-Registry -KeyPath 'HKEY_CURRENT_USER\Printers\Settings' -Pattern * -SearchPropertyName | Export-Csv -Path 'D:\printers.csv' -NoTypeInformation

            or

            Search-Registry -Hive HKEY_CURRENT_USER -KeyPath 'Printers\Settings' -Pattern * -SearchPropertyName | Export-Csv -Path 'D:\printers.csv' -NoTypeInformation

            Searches HKEY_CURRENT_USER (HKCU) on the local computer for printer names and outputs it as a CSV file.
        .EXAMPLE
            Search-Registry -KeyPath 'HKLM:\SOFTWARE\Classes\Installer' -Pattern LastUsedSource -SearchPropertyName -Recurse

            or

            Search-Registry -Hive HKLM -KeyPath 'SOFTWARE\Classes\Installer' -Pattern LastUsedSource -SearchPropertyName -Recurse

            Outputs the LastUsedSource registry entries on the current computer.
        .EXAMPLE
            Search-Registry -KeyPath 'HKCR\.odt' -RegexPattern '.*' -SearchKeyName -MaximumResults 10 -Verbose

            or 

            Search-Registry -Hive HKCR -KeyPath '.odt' -RegexPattern '.*' -SearchKeyName -MaximumResults 10 -Verbose

            Outputs at most ten matches if the specified key exists. 
            This command returns a result if the current computer has a program registered to open files with the .odt extension. 
            The pattern '.*' means match everything.
        .EXAMPLE
            Get-Content Computers.txt | Search-Registry -KeyPath "HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine" -Pattern '*' -SearchPropertyName | Export-Csv -Path 'D:\powershell.csv' -NoTypeInformation

            Searches for any property name in the registry on each computer listed in the file Computers.txt starting at the specified subkey. 
            Output is sent to the specified CSV file.
        .EXAMPLE
            Search-Registry -KeyPath 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace' -SearchPropertyName -Recurse -Verbose

            Searches for the default (nameless) properties in the specified registry key.
        .LINK
            https://stackoverflow.com/questions/54041911/fast-registry-searcher-in-powershell
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByWildCard')]
    Param(
        [Parameter(ValueFromPipeline = $true, Mandatory = $false, Position = 0)]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory = $false, ParameterSetName = 'ByRegex')]
        [string]$RegexPattern,

        [Parameter(Mandatory = $false, ParameterSetName = 'ByWildCard')]
        [string]$Pattern,

        [Parameter(Mandatory = $false)]
        [ValidateSet('HKEY_CLASSES_ROOT','HKEY_CURRENT_CONFIG','HKEY_CURRENT_USER','HKEY_DYN_DATA','HKEY_LOCAL_MACHINE',
                     'HKEY_PERFORMANCE_DATA','HKEY_USERS','HKCR','HKCC','HKCU','HKDD','HKLM','HKPD','HKU')]
        [string]$Hive,

        [string]$KeyPath,
        [int32] $MaximumResults = [int32]::MaxValue,
        [switch]$SearchKeyName,
        [switch]$SearchPropertyName,
        [switch]$SearchPropertyValue,
        [switch]$Recurse
    )
    Begin {
        # detect if the function is called using the pipeline or not
        # see: https://communary.net/2015/01/12/quick-tip-determine-if-input-comes-from-the-pipeline-or-not/
        # and: https://www.petri.com/unraveling-mystery-myinvocation
        [bool]$isPipeLine = $MyInvocation.ExpectingInput

        # sanitize given parameters
        if ([string]::IsNullOrWhiteSpace($ComputerName) -or $ComputerName -eq '.') { $ComputerName = $env:COMPUTERNAME }

        # parse the give KeyPath
        if ($KeyPath -match '^(HK(?:CR|CU|LM|U|PD|CC|DD)|HKEY_[A-Z_]+)[:\\]?') {
            $Hive = $matches[1]
            # remove HKLM, HKEY_CURRENT_USER etc. from the path
            $KeyPath = $KeyPath.Split("\", 2)[1]
        }
        switch($Hive) {
            { @('HKCC', 'HKEY_CURRENT_CONFIG') -contains $_ }   { $objHive = [Microsoft.Win32.RegistryHive]::CurrentConfig;   break }
            { @('HKCR', 'HKEY_CLASSES_ROOT') -contains $_ }     { $objHive = [Microsoft.Win32.RegistryHive]::ClassesRoot;     break }
            { @('HKCU', 'HKEY_CURRENT_USER') -contains $_ }     { $objHive = [Microsoft.Win32.RegistryHive]::CurrentUser;     break }
            { @('HKDD', 'HKEY_DYN_DATA') -contains $_ }         { $objHive = [Microsoft.Win32.RegistryHive]::DynData;         break }
            { @('HKLM', 'HKEY_LOCAL_MACHINE') -contains $_ }    { $objHive = [Microsoft.Win32.RegistryHive]::LocalMachine;    break }
            { @('HKPD', 'HKEY_PERFORMANCE_DATA') -contains $_ } { $objHive = [Microsoft.Win32.RegistryHive]::PerformanceData; break }
            { @('HKU',  'HKEY_USERS') -contains $_ }            { $objHive = [Microsoft.Win32.RegistryHive]::Users;           break }
        }

        # critical: Hive could not be determined
        if (!$objHive) {
            Throw "Parameter 'Hive' not specified or could not be parsed from the 'KeyPath' parameter."
        }

        # critical: no search criteria given
        if (-not ($SearchKeyName -or $SearchPropertyName -or $SearchPropertyValue)) {
            Throw "You must specify at least one of these parameters: 'SearchKeyName', 'SearchPropertyName' or 'SearchPropertyValue'"
        }

        # no patterns given will only work for SearchPropertyName and SearchPropertyValue
        if ([string]::IsNullOrEmpty($RegexPattern) -and [string]::IsNullOrEmpty($Pattern)) {
            if ($SearchKeyName) {
                Write-Warning "Both parameters 'RegexPattern' and 'Pattern' are emtpy strings. Searching for KeyNames will not yield results."
            }
        }

        # create two variables for output purposes
        switch ($objHive.ToString()) {
            'CurrentConfig'   { $hiveShort = 'HKCC'; $hiveName = 'HKEY_CURRENT_CONFIG' }
            'ClassesRoot'     { $hiveShort = 'HKCR'; $hiveName = 'HKEY_CLASSES_ROOT' }
            'CurrentUser'     { $hiveShort = 'HKCU'; $hiveName = 'HKEY_CURRENT_USER' }
            'DynData'         { $hiveShort = 'HKDD'; $hiveName = 'HKEY_DYN_DATA' }
            'LocalMachine'    { $hiveShort = 'HKLM'; $hiveName = 'HKEY_LOCAL_MACHINE' }
            'PerformanceData' { $hiveShort = 'HKPD'; $hiveName = 'HKEY_PERFORMANCE_DATA' }
            'Users'           { $hiveShort = 'HKU' ; $hiveName = 'HKEY_USERS' }
        }

        if ($MaximumResults -le 0) { $MaximumResults = [int32]::MaxValue }
        $script:resultCount = 0
        [bool]$useRegEx = ($PSCmdlet.ParameterSetName -eq 'ByRegex')

        # -------------------------------------------------------------------------------------
        # Nested helper function to (recursively) search the registry
        # -------------------------------------------------------------------------------------
        function _RegSearch([Microsoft.Win32.RegistryKey]$objRootKey, [string]$regPath, [string]$computer) {
            try {
                if ([string]::IsNullOrWhiteSpace($regPath)) {
                    $objSubKey = $objRootKey
                }
                else {
                    $regPath = $regPath.TrimStart("\")
                    $objSubKey = $objRootKey.OpenSubKey($regPath, $false)    # $false --> ReadOnly
                }
            }
            catch {
              Write-Warning ("Error opening $($objRootKey.Name)\$regPath" + "`r`n         " + $_.Exception.Message)
              return
            }
            $subKeys = $objSubKey.GetSubKeyNames()

            # Search for Keyname
            if ($SearchKeyName) {
                foreach ($keyName in $subKeys) {
                    if ($script:resultCount -lt $MaximumResults) {
                        if ($useRegEx) { $isMatch = ($keyName -match $RegexPattern) } 
                        else { $isMatch = ($keyName -like $Pattern) }
                        if ($isMatch) {
                            # for PowerShell < 3.0 use: New-Object -TypeName PSObject -Property @{ ... }
                            [PSCustomObject]@{
                                'ComputerName'     = $computer
                                'Hive'             = $objHive.ToString()
                                'HiveName'         = $hiveName
                                'HiveShortName'    = $hiveShort
                                'Path'             = $objSubKey.Name
                                'SubKey'           = "$regPath\$keyName".TrimStart("\")
                                'ItemType'         = 'RegistryKey'
                                'DataType'         = $null
                                'ValueKind'        = $null
                                'PropertyName'     = $null
                                'PropertyValue'    = $null
                                'PropertyValueRaw' = $null
                            }
                            $script:resultCount++
                        }
                    }
                }
            }

            # search for PropertyName and/or PropertyValue
            if ($SearchPropertyName -or $SearchPropertyValue) {
                foreach ($name in $objSubKey.GetValueNames()) {
                    if ($script:resultCount -lt $MaximumResults) {
                        $data = $objSubKey.GetValue($name)
                        $raw  = $objSubKey.GetValue($name, '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)

                        if ($SearchPropertyName) {
                            if ($useRegEx) { $isMatch = ($name -match $RegexPattern) }
                            else { $isMatch = ($name -like $Pattern) }

                        }
                        else {
                            if ($useRegEx) { $isMatch = ($data -match $RegexPattern -or $raw -match $RegexPattern) } 
                            else { $isMatch = ($data -like $Pattern -or $raw -like $Pattern) }
                        }

                        if ($isMatch) {
                            $kind = $objSubKey.GetValueKind($name).ToString()
                            switch ($kind) {
                                'Binary'       { $dataType = 'REG_BINARY';    break }
                                'DWord'        { $dataType = 'REG_DWORD';     break }
                                'ExpandString' { $dataType = 'REG_EXPAND_SZ'; break }
                                'MultiString'  { $dataType = 'REG_MULTI_SZ';  break }
                                'QWord'        { $dataType = 'REG_QWORD';     break }
                                'String'       { $dataType = 'REG_SZ';        break }
                                default        { $dataType = 'REG_NONE';      break }
                            }
                            # for PowerShell < 3.0 use: New-Object -TypeName PSObject -Property @{ ... }
                            [PSCustomObject]@{
                                'ComputerName'     = $computer
                                'Hive'             = $objHive.ToString()
                                'HiveName'         = $hiveName
                                'HiveShortName'    = $hiveShort
                                'Path'             = $objSubKey.Name
                                'SubKey'           = $regPath.TrimStart("\")
                                'ItemType'         = 'RegistryProperty'
                                'DataType'         = $dataType
                                'ValueKind'        = $kind
                                'PropertyName'     = if ([string]::IsNullOrEmpty($name)) { '(Default)' } else { $name }
                                'PropertyValue'    = $data
                                'PropertyValueRaw' = $raw
                            }
                            $script:resultCount++
                        }
                    }
                }
            }

            # recurse through all subkeys
            if ($Recurse) {
                foreach ($keyName in $subKeys) {
                    if ($script:resultCount -lt $MaximumResults) {
                        $newPath = "$regPath\$keyName"
                        _RegSearch $objRootKey $newPath $computer
                    }
                }
            }

            # close opened subkey
            if (($objSubKey) -and $objSubKey.Name -ne $objRootKey.Name) { $objSubKey.Close() }
        }
    }
    Process{
       if ($isPipeLine) { $ComputerName = @($_) }
       $ComputerName | ForEach-Object {
            Write-Verbose "Searching the registry on computer '$ComputerName'.."
            try {
                $rootKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($objHive, $_)
                _RegSearch $rootKey $KeyPath $_
            }
            catch {
                Write-Error "$($_.Exception.Message)"
            }
            finally {
                if ($rootKey) { $rootKey.Close() }
            }
        }
        Write-Verbose "All Done searching the registry. Found $($script:resultCount) results."
    }
}

##############################
#
# CONVERSIONS, CURRENCY, LENGTH, WEIGHT
#
##############################

# Common Conversions
# Length Conversions
# 1 mile = 5,280 ft = 1,760 yards = 1,609.34 m = 1.609 km
# 1 ft = 12 inches = 0.305 m = 30.48 cm
# 1 inch = 2.54 cm = 25.4 mm
# 1 km = 1,000 m = 0.621 miles
# 1 m = 39.37 inches = .3281 ft = 1.094 yards
# 1 cm = 10 mm = 0.394 inches
# 1 mm = 0.04 inches
# 
# Area Conversions
# 1 section = 1 sq. mile = 640 acres = 2.59 sq km = 259 hectares
# 1 acre = 43,560 sq ft = 0.405 ha = 4047 sq m
# 1 sq ft = 0.093 sq m
# 1 sq km = 1,000,000 sq m = 100 ha = 247.1 acres
# 1 ha = 10,000 sq m = 2.471 ac
# 
# Capacity conversions
# 1 gallon = 4 quarts = 8 pints = 16 cups = 3.785 liters
# 1 quart = 0.946 liters
# 1 liter = 1.057 quarts = 0.264 gallons
# 
# Weight Conversions
# 1 ton = 2,000 lbs = 907.18 kg
# 1 lb = 16 oz = 453.59 g = 0.454 kg
# 1 oz = 28.35 g
# 1 kg = 2.205 lbs = 1,000 g = 1,000,000 mg
# 
# Rate Conversions
# 1 lb/ac = 1.121 kg/ha
# 1 kg/ha = 0.891 lbs/ac
# 1 ppm = 1 mg/kg
# 
# Miscellaneous Conversions
# Degrees Fahrenheit = (9/5 X degrees C) + 32
# Degrees Celsius = 5/9 X (degrees F - 32)
# Area of a Circle = 3.1416 X radius2
# Volume of a cylinder = 3.1416 X radius2X height

function F-toC([double] $fahrenheit) { "$([Math]::Round( (($fahrenheit - 32) / 1.8), 2)) C   [ C = (F-32) * 5 / 9 ]" }
function C-toF([double] $celcius) { "$([Math]::Round( (($celcius * 1.8) + 32), 2)) F   [ F = (9/5 * C) + 32 ]" }
function LB-toKG([double] $lb) { "$($lb * 0.453592) kg   [ 1 kg = 2.205 lbs = 35.2 oz ]" }
function KG-toLB([double] $kg) { "$($kg / 0.453592) lb   [ 1 lb = 0.454 kg = 16 oz ]" }   # update this to split off the ounces and calculate those
Set-Alias F2C F-toC
Set-Alias C2F C-toF

function ConvertTo-Metric
{
    <#
    .Synopsis
        Converts units from imperial to metric
    .Description
        Converts a variety of units from imperial to metric
    .Example
        ConvertTo-Metric 1 pound
    .Link
        https://www.powershellgallery.com/packages/Formulaic/0.2.1.0/Content/ConvertTo-Metric.ps1
    #>
    [OutputType([PSObject])]
    param(
    # The value to convert into metric
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
    [Double]$Value,
    # The unit the value is in
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=1)]
    [ValidateSet('Inch','Inches','Foot','Feet','Yard','Yards','Mile','Miles', 'Pound', 'Pounds', 'Lb', 'Lbs', 'In', 'Mi', 'Ft')]
    [string]$Unit   
    )
        
    process {
        switch ($unit) {
            { 'Inch', 'Inches', 'In' -contains $_ } {                
                New-Object PSObject |
                    Add-Member NoteProperty mm ($value * 25.4) -PassThru |
                    Add-Member NoteProperty cm ($value * 2.54) -PassThru |
                    Add-Member NoteProperty m ($value * .0254) -PassThru |
                    Add-Member NoteProperty km ($value * .0000254) -PassThru 
            }
            { 'Foot', 'Feet', 'Ft' -contains $_ } {
                New-Object PSObject |
                    Add-Member NoteProperty mm ($value * 304.8) -PassThru |
                    Add-Member NoteProperty cm ($value * 30.48) -PassThru |
                    Add-Member NoteProperty m ($value * .3048) -PassThru |
                    Add-Member NoteProperty km ($value * .0003048) -PassThru 
            }
            { 'Yard', 'Yards' -contains $_ } {
                New-Object PSObject |
                    Add-Member NoteProperty km ($value * .0009144) -PassThru |
                    Add-Member NoteProperty m ($value * .9144) -PassThru |
                    Add-Member NoteProperty cm ($value * 91.44) -PassThru |
                    Add-Member NoteProperty mm ($value * 914.4) -PassThru                   
                    
            }
            { 'Mile', 'Miles', 'Mi' -contains $_ } {
                New-Object PSObject |
                    Add-Member NoteProperty km ($value * 1.609) -PassThru |
                    Add-Member NoteProperty m ($value * 1609) -PassThru |
                    Add-Member NoteProperty cm ($value * 160900) -PassThru |
                    Add-Member NoteProperty mm ($value * 1609000) -PassThru                                                           
            }
              
            
            { 'Pound', 'Pound', 'Lbs', 'Lb' -contains $_ }  {
                New-Object PSObject |
                    Add-Member NoteProperty mg ($value * 4536000) -PassThru |                                                         
                    Add-Member NoteProperty g ($value *  4536) -PassThru |
                    Add-Member NoteProperty kg ($value *  .4536 ) -PassThru
            }
            
            { 'Ton', 'Tons', 'Tn' -contains $_ }  {
                New-Object PSObject |
                    Add-Member NoteProperty mg ($value * 900000000) -PassThru |                                                         
                    Add-Member NoteProperty g ($value *  900000) -PassThru |
                    Add-Member NoteProperty kg ($value *  900 ) -PassThru
            }                       
        }
    }
} 

# https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/converting-currencies
# Illustrates how a dynamic parameter can be populated by dynamic data, and how this data is cached
# so IntelliSense won't trigger a new retrieval all the time.
# 100, 66.9 | ConvertTo-Euro -Currency DKK

function ConvertTo-Euro
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [Double] $Value
    )
  
    dynamicparam
    {
        $Bucket = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
    
        $Attributes = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]    
        $AttribParameter = New-Object System.Management.Automation.ParameterAttribute
        $AttribParameter.Mandatory = $true
        $Attributes.Add($AttribParameter)
        
        if ($script:currencies -eq $null)
        {
            $url = 'http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml'
            $result = Invoke-RestMethod  -Uri $url
            $script:currencies = $result.Envelope.Cube.Cube.Cube.currency
        }
        
        $AttribValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($script:currencies)
        $Attributes.Add($AttribValidateSet)
    
        $Parameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter('Currency',[String], $Attributes)
        $Bucket.Add('Currency', $Parameter)
    
        $Bucket
    }
  
    begin
    {
      foreach ($key in $PSBoundParameters.Keys)
      {
          if ($MyInvocation.MyCommand.Parameters.$key.isDynamic)
          {
              Set-Variable -Name $key -Value $PSBoundParameters.$key
          }
      }
    
      $url = 'http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml'
      $rates = Invoke-RestMethod -Uri $url
      $rate = $rates.Envelope.Cube.Cube.Cube | 
      Where-Object { $_.currency -eq $Currency} |
      Select-Object -ExpandProperty Rate
    }
  
    process
    {
        $result = [Ordered]@{
            Value = $Value
            Currency = $Currency
            Rate = $rate
            Euro = ($Value / $rate)
            Date = Get-Date
        }
      
        New-Object -TypeName PSObject -Property $result
    }
}



Function Test-IsWindowsTerminal {
    [cmdletbinding()]
    [Outputtype([Boolean])]

    Param()

    Write-Verbose "Testing processid $pid"

    if ($PSVersionTable.PSVersion.major -ge 6) {
        #PowerShell Core has a Parent property for process objects
        Write-Verbose "Using Get-Process"
        $parent = (Get-Process -id $pid).Parent
        Write-Verbose "Parent process ID is $($parent.id) ($($parent.processname))"
        if ($parent.ProcessName -eq "WindowsTerminal") {
            $True
        }
        Else {
            #check the grandparent process
            $grandparent = (Get-Process -id $parent.id).parent
            Write-Verbose "Grandarent process ID is $($grandparent.id) ($($grandparent.processname))"
            if ($grandparent.processname -eq "WindowsTerminal") {
                $True
            }
            else {
                $False
            }
        }
    } #if Core or later
    else {
        #PowerShell 5.1 needs to use Get-CimInstance
        Write-Verbose "Using Get-CimInstance"

        $current = Get-CimInstance -ClassName win32_process -filter "processid=$pid"
        $parent = Get-Process -id $current.parentprocessID
        Write-Verbose "Parent process ID is $($parent.id) ($($parent.processname))"
        if ($parent.ProcessName -eq "WindowsTerminal") {
            $True
        }
        Else {
            #check the grandparent process
            $cimGrandparent = Get-CimInstance -classname win32_process -filter "Processid=$($parent.id)"
            $grandparent = Get-Process -id $cimGrandparent.parentProcessId
            Write-Verbose "Grandarent process ID is $($grandparent.id) ($($grandparent.processname))"
            if ($grandparent.processname -eq "WindowsTerminal") {
                $True
            }
            else {
                $False
            }
        }
    } #PowerShell 5.1
} #close function

function color {
    # color "cyan" "black" will set background colour to cyan and foreground (text) color to black
    # Powershell color names are:
    #    Black White
    #    Gray DarkGray
    #    Red DarkRed
    #    Blue DarkBlue
    #    Green DarkGreen
    #    Yellow DarkYellow
    #    Cyan DarkCyan
    #    Magenta DarkMagenta

    [CmdletBinding()]
    param( [Parameter(Mandatory)] [string] $BackgroundColor, 
           [Parameter(Mandatory)] [string] $ForegroundColor )

    $a = (Get-Host).UI.RawUI
    $a.BackgroundColor = $BackgroundColor
    $a.ForegroundColor = $ForegroundColor
    cls
}

function SwapTitleAuthor($separator) {
    # if no separator defined, use " - "
    if ($null -eq $separator) { $separator = " - " }
    # split name into $part1 $part2 $extn => $part2 $separator $part1 $extn
    echo $separator
}

function Convert-BooksToMobi($folder) {
    function color ($bc, $fc) {
        # https://social.technet.microsoft.com/Forums/ie/en-US/4b43f071-abf5-4a65-9048-82d474473a8e/how-can-i-set-the-powershell-console-background-color-not-the-text-background-color?forum=winserverpowershell
        $a = (Get-Host).UI.RawUI
        $a.BackgroundColor = $bc
        $a.ForegroundColor = $fc
        Clear-Host   # cls
    }

    $ebookconvert = "C:\Program Files\Calibre2\ebook-convert.exe"
    if (!(Test-Path $ebookconvert)) {
        ""
        "Could not find the Calibre ebook-convert.exe tool at:"
        "   C:\Program Files\Calibre2\ebook-convert.exe"
        "This can be fixed by:"
        "   choco inst calibre -y"
        ""
        break
    }
    color "cyan" "black"   # force this as ebook-convert.exe always uses unreadable black text(!)

    # epub, azw3, docx, doc, odt, pdf, txt
    # .cbc Comic book collections
    # https://manual.calibre-ebook.com/conversion.html
    foreach ($i in gci $folder -File) {
        if ($i.Extension -in ".epub", ".azw3", ".docx", ".doc", ".odt", ".pdf", ".txt") {
            $parent = Split-Path $i.FullName
            $parentname = Split-Path $parent -Leaf
            $extension = ($i.Extension)
            $extfolder = "$($parent)\$($parentname)$($extension)"
            $mobiname = "$parent\$($i.BaseName).mobi"
            $cannotconvert = "$parent\$($i.BaseName).cannotconvert"
            $src = $i.FullName
            $dst = $mobiname
            ""
            ""
            ""
            "####################"
            "#"
            "# $parent"
            "# $parentname"
            "# Name : $($i.Name)"
            "# Extn : $($i.Extension)"
            "# Src  : $src"
            "# Dst  : $dst"
            "#"
            "####################"
            ""
            $cmd = "& '$ebookconvert' '$src' '$dst'"
            iex $cmd   # & will run $convert-exe as an .exe https://stackoverflow.com/questions/3592851/executing-a-command-stored-in-a-variable-from-powershell
            if (!(Test-Path $dst)) {
                # Test if the .mobi was *not* created
                md $cannotconvert
                move $i.FullName $cannotconvert\$i
            }
            else {
                # If the .mobi was created, move the old file in there
                if (!(Test-Path $extfolder)) {
                    md $extfolder
                }
                move $i.FullName $extfolder\$i
            }
        }
    }
    $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkBlue')   # Odd syntax, but this is perfectly valid to define $bckgrnd
    $Host.UI.RawUI.ForegroundColor = 'White'
    $Host.PrivateData.ErrorForegroundColor = 'Red'
    $Host.PrivateData.ErrorBackgroundColor = $bckgrnd
    $Host.PrivateData.WarningForegroundColor = 'Magenta'
    $Host.PrivateData.WarningBackgroundColor = $bckgrnd
    $Host.PrivateData.DebugForegroundColor = 'Yellow'
    $Host.PrivateData.DebugBackgroundColor = $bckgrnd
    $Host.PrivateData.VerboseForegroundColor = 'Green'
    $Host.PrivateData.VerboseBackgroundColor = $bckgrnd
    $Host.PrivateData.ProgressForegroundColor = 'Cyan'
    $Host.PrivateData.ProgressBackgroundColor = $bckgrnd
    "Use 'Clear-Host' to reset console values."
}

function Get-RegistryKeyPropertiesAndValues {
    # Template fuction, to demonstrate, need to build up functions to write etc, basic registry manipulation tools
    # Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Microsoft Antimalware\Exclusions\Paths' -Name 'TheExcludedPath' -Value 0 -Type DWord -Force
    # ToDo: Get-RegistryProperty, Set-RegistryProperty, Get-RegistryKeyAsHashTable, etc
    # ToDo: All of the same, but for .ini files
    Push-Location
    Set-Location 'HKCU:\Environment'
    Get-Item . | Select-Object -ExpandProperty Property | ForEach-Object {
        New-Object psobject -Property @{"Property" = $_ ; "Value" = (Get-ItemProperty -Path . -Name $_).$_}
    } | Format-Table property, value -AutoSize
    Pop-Location
}

# These are just examples of constructing date-time strings for filenames etc
function DateTimeNowString { Get-Date -format "yyyy-MM-dd__HH-mm-ss" }   # HH is 24-hour format, hh is 12-hour format
function DateTimeNowStringMinutes { Get-Date -format "yyyy-MM-dd__HH-mm" }   # Use only minute granularity

# Equivalent to unix 'touch', this will update the LastWriteTime of a file to current datetime.
function Update-LastWriteTime ($file)
{
    if ($file -eq $null) { throw "No filename supplied" }
    if (Test-Path $file) { (Get-ChildItem $file).LastWriteTime = Get-Date }
    else { New-Item -ItemType File $file }
}
Set-Alias -Name touch -Value Update-LastWriteTime

function Install-ModuleToDirectory {
    [CmdletBinding()] [OutputType('System.Management.Automation.PSModuleInfo')]
    param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]                                    $Name,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidateScript({ Test-Path $_ })] $Destination
    )

    if (($Profile -like "\\*") -and (Test-Path (Join-Path $UserModulePath $Name))) {
        if (Test-Administrator -eq $true) {
            # Nothing in here will happen unless working on laptop with a network share
            Uninstall-Module $Name -Force -Verbose
            # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!) so use Save-Module
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
        }
        else {
            "Module found on network share module path but need to be administrator and connected to VPN"
            "to correctly move Modules into the users module folder on C:\"
            pause
        }
    }
    elseif (($Profile -like "\\*") -and (Test-Path (Join-Path $Profile $Name))) {
        if (Test-Administrator -eq $true) {
            # Nothing in here will happen unless working on laptop with a network share
            Uninstall-Module $Name -Force -Verbose
            # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!) so use Save-Module
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
        }
        else {
            "Module found on network share module path but need to be administrator and connected to VPN"
            "to correctly move Modules into the users module folder on C:\"
            pause
        }
    }
    elseif (Test-Path (Join-Path $AdminModulePath $Name)) {
        if (Test-Administrator -eq $true) {
            # Nothing in here will happen unless working on laptop with a network share
            Uninstall-Module $Name -Force -Verbose
            # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!) so use Save-Module
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
        }
        else {
            "Module found on in Admin Modules folder: C:\Program Files\WindowsPowerShell\Modules."
            "Need to be Admin to correctly move Modules into the users module folder on C:\"
            pause
        }
    }
    else {
        Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
        Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
    }
    $out = ""; foreach ($i in (Get-Command -Module $Name).Name) {$out = "$out, $i"} ; "" ; Write-Wrap $x.trimstart(", ") ; ""
    # return (Get-Module)
}

# Install-ModuleToDirectory -Name 'XXX' -Destination 'E:\Modules'
# try {
#     # Note additional switches if required: -Repository $MyRepoName -Credential $Credential
#     # If the module is already installed, use Update, otherwise use Install
#     if ([bool](Get-Module $Name -ListAvailable)) {
#          Update-Module $Name -Verbose -ErrorAction Stop 
#     } else {
#          Install-Module $Name -Scope CurrentUser -Verbose -ErrorAction Stop
#     }
# } catch {
#     # But if something went wrong, just -Force it, hard.
#     Install-Module $Name -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
# }


# $ModuleNetShare = "$(Split-Path $ProfileNetShare)\Modules\$MyModule"
# $ModuleCProfile = "C:\Users\$env:Username\Documents\WindowsPowerShell\Modules\$MyModule"
# $ModuleNetShare
# $ModuleCProfile

#     $success = 0
#     # if ($null -ne $(Test-Path $ModuleNetShare)) {
#     # Only run this if $Profile is pointing at Net Share
#     # Note that the uninstalls will fail unless connected to the VPN!
#     if ((Test-Path $ModuleNetShare) -and ($Profile -like "\\*")) {
#         if (Test-Administrator -eq $true) {
#             # Nothing in here will happen unless working on laptop with a network share
#             # First uninstall, then reinstall to get latest version, then move it to $Profile
#             Uninstall-Module $MyModule -Force -Verbose
#             Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!)
#             Move-Item $ModuleNetShare $ModuleCProfile -Force -Verbose
#             Uninstall-Module $MyModule -Force -Verbose                    # Need to uninstall again to clear network share reference from registry
#             Import-Module $MyModule -Scope Local -Force -Verbose          # Finally, import the version in C:\
#             $success = 1
#         }
#         else {
#             "Module found on network share module path but need to be administrator and connected to VPN"
#             "to correctly move Modules into the users module folder on C:\"
#             pause
#         }
#     }
#     if ((Test-Path "C:\Program Files\WindowsPowerShell\Modules\$MyModule") -and (Test-Administrator -eq $true)) {
#         # This is if the module has been loaded into the Administrator folder.
#         # This will move it to the user folder and update
#         Uninstall-Module $MyModule -Force -Verbose
#         Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!), so same situation as before
#         Move-Item $ModuleNetShare $ModuleCProfile -Force -Verbose
#         Uninstall-Module $MyModule -Force -Verbose                    # Need to uninstall again to clear network share reference from registry
#         Import-Module $MyModule -Scope Local -Force -Verbose          # Finally, import the version in C:\
#         $success = 1
#     }
#     else {
#         "Module $MyModule found in 'C:\Program Files\WindowsPowerShell\Modules' but need to be administrator"
#         "to correctly move Modules into the users module folder on C:\"
#         pause
#     }
# 
#     if ($success -eq 0) {
#         try {
#             # Note additional switches if required: -Repository $MyRepoName -Credential $Credential
#             # If the module is already installed, use Update, otherwise use Install
#             if ([bool](Get-Module $MyModule -ListAvailable)) {
#                  Update-Module $MyModule -Verbose -ErrorAction Stop 
#             } else {
#                  Install-Module $MyModule -Scope CurrentUser -Verbose -ErrorAction Stop
#             }
#         } catch {
#             # But if something went wrong, just -Force it, hard.
#             Install-Module $MyModule -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
#         }
#     }
# }

function Install-ProfileForceLocalForFasterLoading {
    # To get around the VPN console load time issue when opening PowerShell consoles while
    # connected to a network share on a corporate VPN. i.e. ING(!)

    $ProfileLeaf = split-path $env:userprofile -leaf   # Just get the correct username in spite of any changes to username!
    $OSver = (Get-WMIObject win32_operatingsystem).Name
    $PSver = $PSVersionTable.PSVersion.Major

    function Test-Administrator {
        $user = [Security.Principal.WindowsIdentity]::GetCurrent();
        (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    if (Test-Administrator -eq $true -and $OSver -like "*Enterprise*") {
        Write-Host ""
        Write-Host "`n========================================" -ForegroundColor Green
        Write-Host 'Configure $Profile to stop calls to network share over VPN.' -ForegroundColor Green
        Write-Host ''
        Write-Host "This section is experimental, to provide a way to override corporate VPN functionality" -F Yellow -B Black
        Write-Host "By default, this will be skipped. This is useful on corporate laptops when working remotely." -F Yellow -B Black
        Write-Host "This section will only show if you are running as Administrator." -F Yellow -B Black
        Write-Host "========================================`n" -ForegroundColor Green
        $check =  "Do you want to update `$Profile.AllUsersAllHosts to always redirect to C:\Users\$($env:Username) ? "
        $check += "This will significantly speed up PowerShell load times when working with a network share over a VPN. "
        $check += "This option is usually only useful when working remotely on a laptop and connected over a VPN. "
        Write-Wrap $check
        Write-Host "`$Profile.AllUsersAllHosts = C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1" -F Cyan -B Black
        Write-Host "`$Profile.CurrentHostCurrentUser = C:\Users\$($env:Username)\Documents\WindowsPowerShell\$(($Profile).split('\')[-1])_profile.ps1" -F Cyan -B Black
        Write-Host ""
        $ProfileAdmin = "C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1"
        # $ProfileName = (($).split("\"))[-1]   # Not needed, but Do this to capture correct prefix PowerShell, VSCode, etc
        $ProfileUser = "C:\Users\$ProfileLeaf\Documents\WindowsPowerShell\$(($Profile).split('\')[-1])"
    
        # Note: Cannot rely upon these variables if BeginSystemConfig is run inside an existing session
        #       as the original $profile has already been overwritten so test them to see if empty after creation.
        Write-Host "Press 'X' to create and use the VPN bypass." -F Yellow -B Black
        Write-Host "Press any other key to skip this function." -F Yellow -B Black
    
        $x = New-Object System.Management.Automation.Host.ChoiceDescription "&X", "X";
        $no  = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No";
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($x, $no);
        $caption = ""   # Did not need this before, but now getting odd errors without it.
        $answer = $host.ui.PromptForChoice($caption, $message, $choices, 1)   # Set to 0 to default to "yes" and 1 to default to "no"
    
        if ($answer -eq 0) {
            # 1. Create $Profile.AllUsersAllHosts
            # 2. Add: Set-Location C:\Users\$($env:Username)   (override Admin defaulting into system32)
            # 3. Add: dotsource C:\Users\$($env:Username)\Documents\WindowsPowerShell\$($ShellId)_profile.ps1
            #         As this is the normal offline location, 
            # 4. Add: 
            # 5. mklink to somewhere easier like C:\PS and add scripts folder to path?
    
            if (!(Test-Path $(Split-Path $ProfileAdmin))) { mkdir $(Split-Path $ProfileAdmin) -Force }
            if (!(Test-Path $(Split-Path $ProfileUser))) { mkdir $(Split-Path $ProfileUser) -Force }
    
            $ProfileUserCFormatted = '"C:\Users\$($env:Username)\Documents\WindowsPowerShell\$(($Profile).split(`"\`")[-1])"'
    
            Write-Host "`nCreating backup of existing profile ..."
            if (Test-Path $ProfileAdmin) { Move-Item -Path "$($ProfileAdmin)" -Destination "$($ProfileAdmin)_$(Get-Date -format "yyyy-MM-dd__HH-mm-ss").txt" }
    
            Write-Host ":: Writing VPN bypass to $ProfileAdmin" -F Yellow -B Black
            Write-Host ""
            Set-Content -Path $ProfileAdmin -Value "Set-Location C:\Users\`$(`$env:Username)   # Default to user folder, even if start as Admin" -PassThru
            Add-Content -Path $ProfileAdmin -Value "" -PassThru
            Add-Content -Path $ProfileAdmin -Value "# Need to save NetShare locations as system will attempt these at times:" -PassThru
            Add-Content -Path $ProfileAdmin -Value "`$ProfileNetShare = `$Profile.CurrentUserCurrentHost" -PassThru
            Add-Content -Path $ProfileAdmin -Value "`$ProfileNetShareAllHosts = `$Profile.CurrentUserAllHosts" -PassThru
            Add-Content -Path $ProfileAdmin -Value "" -PassThru
            Add-Content -Path $ProfileAdmin -Value "# Prevent loading profile from network share" -PassThru
            Add-Content -Path $ProfileAdmin -Value "if (`$Profile.CurrentUserCurrentHost -ne $ProfileUserCFormatted) {" -PassThru
            Add-Content -Path $ProfileAdmin -Value "    `$Profile = $ProfileUserCFormatted" -PassThru
            Add-Content -Path $ProfileAdmin -Value "    if (Test-Path `"`$Profile`") { . `$Profile }" -PassThru
            Add-Content -Path $ProfileAdmin -Value "}" -PassThru
            Write-Host ""
            Write-Host ":: Admin profile updated" -F Yellow -B Black
            Write-Host ""
        }
        else {
            Write-Host ""
            Write-Host "Skipping VPN profile bypass. This is due to slow load tiems when profile is loaded from a network share." -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "To install this improvement (which can increase console / script startup times from 6-9 sec to 1 sec," -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "run this script as Administrator. Note that this also checks that the OS is Windows 'Enterprise', this is" -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "done to ensure that it is not loaded on Windows Server environments or home desktops running Windows Professional." -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "'Get-ExecutionPolicy -List' (show current execution policy list):`n"
            Get-ExecutionPolicy -List | ft
            Write-Host ""
        }
    }
}    



function Edit-Custom ($Name) {
    if ($name -eq "Tools") { $Title = "Tools" ; $Name = "Roy" ; $GistUser = "roysubs" ; $GistId = "5c6a16ea0964cf6d8c1f9eed7103aec8" }
    if ($name -eq "Custom-Tools") { $Title = "Tools" ; $Name = "Roy" ; $GistUser = "roysubs" ; $GistId = "5c6a16ea0964cf6d8c1f9eed7103aec8" }
    if ($name -eq "Roy") { $Title = "Tools" ; $Name = "Roy" ; $GistUser = "roysubs" ; $GistId = "5c6a16ea0964cf6d8c1f9eed7103aec8" }
    if ($name -eq "Edwin") { $Title = "Edwin" ; $Name = "Edwin" ; $GistUser = "e-d-h" ; $GistId = "fd6e178848214614e373c5c36410f648" }
    if ($name -eq "Kevin") { $Title = "Kevin" ; $Name = "Kevin" ; $GistUser = "????" ; $GistId = "????" }

    $CustomTools = "C:\Users\$env:USERNAME\Documents\WindowsPowerShell\Modules\Custom-$Title\Custom-$Title.psm1"
    if (!(Test-Path $CustomTools)) { "`nModule is not present`n$CustomTools`n" ; break }   # New-Item -ItemType File $CustomTools
    code $CustomTools
}

function PullCustomFromGist ($Name) {
    if ($name -eq "Roy") { $Title = "Tools" ; $Name = "Roy" ; $GistUser = "roysubs" ; $GistId = "5c6a16ea0964cf6d8c1f9eed7103aec8" }
    if ($name -eq "Edwin") { $Title = "Edwin" ; $Name = "Edwin" ; $GistUser = "e-d-h" ; $GistId = "fd6e178848214614e373c5c36410f648" }
    if ($name -eq "Kevin") { $Title = "Kevin" ; $Name = "Kevin" ; $GistUser = "????" ; $GistId = "????" }
    $UserProfileLeaf = Split-Path $env:USERPROFILE -Leaf

    Write-Host ""
    Write-Host "Downaload and install latest Custom-$Title.psm1 Module from Gist." -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "Note: Will first backup and datetime stamp current Custom-$Title.psm1 in the Mdoules parent folder." -ForegroundColor Yellow -BackgroundColor Black
    
    $ModuleRoot = "C:\Users\$UserProfileLeaf\Documents\WindowsPowerShell\Modules\Custom-$Title"
    if (-not (Test-Path $ModuleRoot)) { md $ModuleRoot }
    # if (!(Test-Path $ModuleRoot)) { New-Item -Type Directory $ModuleRoot -Force }
    $ModulePSM1 = Join-Path $ModuleRoot "Custom-$Title.psm1"

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls } catch { }   # Windows 7 compatible
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }   # Windows 10 compatible
    Clear-DnsClientCache
    [System.Net.ServicePointManager]::DnsRefreshTimeout = 0;
    
    $now = Get-Date -format "yyyy-MM-dd__HH-mm-ss"
    if (Test-Path $ModuleRoot) {
        if (Test-Path $ModulePSM1) { cp $ModulePSM1 "$(Split-Path $ModuleRoot)\Custom-$($Title)_$($now).psm1" }

        iwr "https://gist.github.com/$GistUser/$GistId/raw" -Headers @{"Cache-Control"="no-cache"} | select -expand content | Out-File $ModulePSM1
        if (Test-Path $ModulePSM1) {
            if (![bool](Get-Module $ModulePSM1 -ListAvailable)) {
                Import-Module -FullyQualifiedName $ModulePSM1 -Force -Verbose
            }
        }
    }
}

function PushCustomToGist ($Name) {
    if ($name -eq "Roy") { $Title = "Tools" ; $Name = "Roy" ; $GistUser = "roysubs" ; $GistId = "5c6a16ea0964cf6d8c1f9eed7103aec8" }
    if ($name -eq "Edwin") { $Title = "Edwin" ; $Name = "Edwin" ; $GistUser = "e-d-h" ; $GistId = "fd6e178848214614e373c5c36410f648" }
    if ($name -eq "Kevin") { $Title = "Kevin" ; $Name = "Kevin" ; $GistUser = "????" ; $GistId = "????" }
    
    Write-Host ""
    Write-Host "Upload Custom-$Title.psm1 Module from Modules folder to Gist." -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "Note: Only $Name can upload to his Gist using his GitHub key and password." -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "Note: Will first backup and datetime stamp current Custom-$Title.psm1 in the Mdoules parent folder." -ForegroundColor Yellow -BackgroundColor Black
    
    $ModuleRoot = "C:\Users\$($env:USERNAME)\Documents\WindowsPowerShell\Modules\Custom-$Title"
    if (-not (Test-Path $ModuleRoot)) { md $ModuleRoot }
    $ModulePSM1 = Join-Path $ModuleRoot "Custom-$Title.psm1"
    $GistPassword = "$env:TEMP\ps_Gist-Secure-Password-$Name.txt"

    "Uploading $ModulePSM1 to $Name GitHub"
    "GitHub Gist details: $GistUser $GistId"

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls } catch { }   # Windows 7 compatible
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }   # Windows 10 compatible
    Clear-DnsClientCache
    [System.Net.ServicePointManager]::DnsRefreshTimeout = 0
    
    $now = Get-Date -format "yyyy-MM-dd__HH-mm-ss"
    if (Test-Path $ModuleRoot) {
        if (Test-Path $ModulePSM1) { cp $ModulePSM1 "$(Split-Path $ModuleRoot)\Custom-$($Title)_$($now).psm1" }

        if (!(Test-Path $GistPassword)) {
            Read-Host "Enter the GitHub password for $GistUser" -AsSecureString | ConvertFrom-SecureString | Out-File $GistPassword
            Write-Host "Password now saved securely to `$env:TEMP\ps_Roy-Gist-Secure-Password.txt`nIf you want to regenerate a new password, delete that file."
        }
        $password = Get-Content $GistPassword | ConvertTo-SecureString
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $GistUser, $password

        # https://adamtheautomator.com/powershell-get-credential/
        # https://stackoverflow.com/questions/6239647/using-powershell-credentials-without-being-prompted-for-a-password
        
        try {
            echo "iwr `"https://gist.github.com/$($GistUser)/$($GistId)/raw`""
            iwr "https://gist.github.com/$GistUser/$GistId/raw"
        }
        catch {
            Write-Host "Invoke-WebRequest failed. You might be behind a firewall / VPN or`nthe Internet Explorer engine might not be fully initialised.`nTrying again by forcing Internet Explorer to start now."
            Start-Process -WindowStyle Hidden "C:\Program Files\Internet Explorer\iexplore.exe" "www.google.com"
            # http://blogs.msdn.com/powershell/archive/2006/09/10/controlling-internet-explorer-object-from-powershell.aspx
            # https://social.technet.microsoft.com/Forums/ie/en-US/e54555bd-00bb-4ef9-9cb0-177644ba19e2/how-to-open-url-through-powershell
            try {
                iwr "https://gist.github.com/$GistUser/$GistId/raw"
            }
            catch {
                throw "Second Invoke-WebRequest attempt also failed. Makes sure that Internet Explorer is fully initialised and check internet / VPN."
                Start-Process -WindowStyle Hidden "C:\Program Files\Internet Explorer\iexplore.exe" "www.google.com"
                rm $GistPassword -Force -EA Silent
            }
        }
        
        # Test if Posh-Gist is installed
        if (![bool](Get-Module Posh-Gist -ListAvailable)) { "Posh-Gist must be installed to run this" ; break }
        
        # Securely upload the Gists, testing existence and non-zero size
        function Push-GistUpdate ($file, $id) {
            if (Test-Path($file)) {
                if ((Get-Item $file).length -gt 0kb) {
                    Update-Gist -Credential $cred -Id $id -Update $file
                }
            }
        }
        Push-GistUpdate $ModulePSM1 $GistId
        
        # Try clearing the DNS client cache to resolve the endpoint caching (this does not work in my experience)
        Clear-DnsClientCache
        [System.Net.ServicePointManager]::DnsRefreshTimeout = 0
    }
}

function Enable-Extensions {
    if (Test-Path "$($Profile)_extensions.ps1") {
        if ($MyInvocation.InvocationName -eq "Enable-Extensions") {
            "`nWarning: Must dotsource Enable-Extensions or it will not be added!`n`n. Enable-Extensions`n"
        } else {
            if (Test-Path "$($Profile)_extensions.ps1") { . "$($Profile)_extensions.ps1" }
            else { "`nProfile Extensions not found" }
        }
    }
}

Function Read-MyHost {
    # https://www.petri.com/prompt-answers-powershell
    # https://www.petri.com/building-a-powershell-console-menu-revisited-part-1
    [cmdletbinding()]
    Param(
    [Parameter(Position=0,Mandatory,HelpMessage="Enter the message prompt.")]
    [ValidateNotNullorEmpty()]
    [string]$Message,
    [Parameter(Position=1,Mandatory,HelpMessage="Enter key property name or names separated by commas.")]
    [System.Management.Automation.Host.FieldDescription []]$Key,
    [Parameter(HelpMessage = "Text to display as a title for the prompt.")]
    [string]$PromptTitle = "",
    [Parameter(HelpMessage = "Convert the result to an object.")]
    [switch]$AsObject
    )
     
    $response = $host.ui.Prompt($PromptTitle,$Message,$Key)
     
    if ($AsObject) {
        #create a custom object
        New-Object -TypeName PSObject -Property $response
    }
    else {
        #write the result to the pipeline
        $response
    }
     
} #end function

#An alternative to the built-in PromptForChoice providing a consistent UI across different hosts
function Get-Choice {
    # https://powershellone.wordpress.com/2015/09/10/a-nicer-promptforchoice-for-the-powershell-console-host/
    # Get-Choice "Pick Something!" (echo Option1 Option2 Option3) 2

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,Position=0)]
        $Title,

        [Parameter(Mandatory=$true,Position=1)]
        [String[]]
        $Options,

        [Parameter(Position=2)]
        $DefaultChoice = -1
    )
    if ($DefaultChoice -ne -1 -and ($DefaultChoice -gt $Options.Count -or $DefaultChoice -lt 1)){
        Write-Warning "DefaultChoice needs to be a value between 1 and $($Options.Count) or -1 (for none)"
        exit
    }
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $script:result = ""
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog
    $form.BackColor = [Drawing.Color]::White
    $form.TopMost = $True
    $form.Text = $Title
    $form.ControlBox = $False
    $form.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    #calculate width required based on longest option text and form title
    $minFormWidth = 100
    $formHeight = 44
    $minButtonWidth = 70
    $buttonHeight = 23
    $buttonY = 12
    $spacing = 10
    $buttonWidth = [Windows.Forms.TextRenderer]::MeasureText((($Options | sort Length)[-1]),$form.Font).Width + 1
    $buttonWidth = [Math]::Max($minButtonWidth, $buttonWidth)
    $formWidth =  [Windows.Forms.TextRenderer]::MeasureText($Title,$form.Font).Width
    $spaceWidth = ($options.Count+1) * $spacing
    $formWidth = ($formWidth, $minFormWidth, ($buttonWidth * $Options.Count + $spaceWidth) | Measure-Object -Maximum).Maximum
    $form.ClientSize = New-Object System.Drawing.Size($formWidth,$formHeight)
    $index = 0
    #create the buttons dynamically based on the options
    foreach ($option in $Options){
        Set-Variable "button$index" -Value (New-Object System.Windows.Forms.Button)
        $temp = Get-Variable "button$index" -ValueOnly
        $temp.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
        $temp.UseVisualStyleBackColor = $True
        $temp.Text = $option
        $buttonX = ($index + 1) * $spacing + $index * $buttonWidth
        $temp.Add_Click({ 
            $script:result = $this.Text; $form.Close() 
        })
        $temp.Location = New-Object System.Drawing.Point($buttonX,$buttonY)
        $form.Controls.Add($temp)
        $index++
    }
    $shownString = '$this.Activate();'
    if ($DefaultChoice -ne -1){
        $shownString += '(Get-Variable "button$($DefaultChoice-1)" -ValueOnly).Focus()'
    }
    $shownSB = [ScriptBlock]::Create($shownString)
    $form.Add_Shown($shownSB)
    [void]$form.ShowDialog()
    $result
}

function Move-ScrollLock {
    Clear-Host
    $timeout = 59
    "Active Keep-alive: ScrollLock will be hit every $timeout seconds..."
    $WshShell = New-Object -Com "WScript.Shell"

    while ($true) {
        $WshShell.SendKeys("{SCROLLLOCK}")
        Sleep -Milliseconds 100
        $WshShell.SendKeys("{SCROLLLOCK}")
        Sleep -Seconds $timeout
    }
}
# Test 1: PS Window open and active in foreground.                      => Stays open forever, no timeout.
# Test 2: PS Window open and active in foreground, RDP window minimized => Times out as normal.
# Test 3: PS Window open and active in foreground, RDP not minimised    => ?

function Start-KeepAliveEdwin {
    # Background job to hit ScrollLock every 3 minutes
    Start-Job {
        $wsh = New-Object -ComObject WScript.Shell
        while($true) {
            $wsh.SendKeys('{SCROLLLOCK}')
            Start-Sleep 180
        }
    }
 
    # # Runspace alternative
    # $rs = [runspacefactory]::CreateRunspace([initialsessionstate]::CreateDefault2())
    # $rs.Open()
    # 
    # # Create a [powershell] object to bind our script to the above runspaces
    # $ps = [powershell]::Create().AddScript({
    #     $wsh = New-Object -ComObject WScript.Shell
    #     while($true) {
    #         $wsh.SendKeys('{SCROLLLOCK}')
    #         Start-Sleep 60
    #     }
    # })
    # 
    # # Tell powershell to run the code in the runspace we created
    # $ps.Runspace = $rs
    # 
    # # Invoke the code asynchronously
    # $handle = $ps.BeginInvoke()
}

function Show-Tools {
    # Use with Show-Tools | sls <searchterm>
    cat "$HomeFix\Documents\WindowsPowerShell\Modules\Custom-Tools\Custom-Tools.psm1"
}

Function Set-DefaultDownloadPath
{
    Param ( [String] $DownloadPath = "%USERPROFILE%\Downloads" )
	
    $userShellFoldersPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    if((Test-Path -Path $DownloadPath) -eq $false) {
         New-Item $DownloadPath -Type Directory -ErrorAction Stop | Out-Null
    }
    if((Get-ItemProperty $userShellFoldersPath).'{374DE290-123F-4565-9164-39C4925E467B}')
    {
        Set-ItemProperty -Path $userShellFoldersPath -Name '{374DE290-123F-4565-9164-39C4925E467B}' -Value $DownloadPath
    }
    #Windows 10
    if((Get-ItemProperty $userShellFoldersPath).'{7D83EE9B-2244-4E70-B1F5-5393042AF1E4}')
    {
        Set-ItemProperty -Path $userShellFoldersPath -Name '{7D83EE9B-2244-4E70-B1F5-5393042AF1E4}' -Value $DownloadPath
    }
    
	#Restart Explorer to change it immediately   
	Stop-Process -Name Explorer
}

Function ChangeWinDefaultDownloadPath
{
    Param ( [String] $DownloadPath = "%USERPROFILE%\Downloads" )
	
    $userShellFoldersPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    if((Test-Path -Path $DownloadPath) -eq $false) {
         New-Item $DownloadPath -Type Directory -ErrorAction Stop | Out-Null
    }
    if((Get-ItemProperty $userShellFoldersPath).'{374DE290-123F-4565-9164-39C4925E467B}')
    {
        Set-ItemProperty -Path $userShellFoldersPath -Name '{374DE290-123F-4565-9164-39C4925E467B}' -Value $DownloadPath
    }
    #Windows 10
    if((Get-ItemProperty $userShellFoldersPath).'{7D83EE9B-2244-4E70-B1F5-5393042AF1E4}')
    {
        Set-ItemProperty -Path $userShellFoldersPath -Name '{7D83EE9B-2244-4E70-B1F5-5393042AF1E4}' -Value $DownloadPath
    }
    
	#Restart Explorer to change it immediately   
	Stop-Process -Name Explorer
}

function Create-ScheduledTasks {
    # Object is to create a task that runs at logon with elevated privileges.
    # This is not possible with a startup folder or startup registry task (specifically blocked by Microsoft by security)
    # but is possible by creating a Scheduled Task (though a bit trickier to setup).

    # Becuase the SysTray task is running elevated, starting tasks from here bypasses UAC (User Access Control) popups.    
    # These tasks trigger at logon with credentials of user running this script.
    # -RunLevel Highest is the key part that starts the task with elevated privileges.
    
    $action = New-ScheduledTaskAction -Execute "powershell.exe"
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -RunLevel Highest -UserId (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName)
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal
    Register-ScheduledTask PowerShellAdminAtLogon -InputObject $task

    # $action = New-ScheduledTaskAction -Execute "C:\0\SysTray\SysTray.lnk"
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-NoProfile -WindowStyle Hidden -File "C:\0\SysTray\SysTray.ps1"'
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -RunLevel Highest -UserId (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName)
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal
    Register-ScheduledTask SysTrayAdminAtLogon -InputObject $task

    # $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -WindowStyle Hidden -command "& {get-eventlog -logname Application -After ((get-date).AddDays(-1)) | Export-Csv -Path C:\0\applog.csv -Force -NoTypeInformation}"'
}
function Start-KeepAlive {
    <#
    .Synopsis
       This is a pecking bird function, a press on the <Ctrl> key will run every 5 minues.
    .DESCRIPTION
       This function will run a background job to keep your computer alive. By default a KeyPess of the <Ctrl> key will be pushed every 3 minutes.
       Please be aware that this is a short term workaround to allow you to complete an otherwise impossible task, such as download a large file.
       This function should only be run when your computer is locked in a secure location.
    .EXAMPLE
       Start-KeepAlive
       Id     Name            PSJobTypeName   State         HasMoreData     Location            
       --     ----            -------------   -----         -----------     --------            
       90     KeepAlive       BackgroundJob   Running       True            localhost           
    
       KeepAlive set to run until 10/01/2012 00:35:03
    
       By default the keepalive will run for 1 hour, with a keypress every 5 minutes.
    .EXAMPLE
       Start-KeepAlive -KeepAliveHours 3
       Id     Name            PSJobTypeName   State         HasMoreData     Location            
       --     ----            -------------   -----         -----------     --------            
       92     KeepAlive       BackgroundJob   Running       True            localhost           
    
       KeepAlive set to run until 10/01/2012 02:36:12
       
       You can specify a longer KeepAlive period using the KeepAlive parameter E.g. specify 3 hours
    .EXAMPLE
       Start-KeepAlive -KeepAliveHours 2 -SleepSeconds 600
       
       You can also change the default period between each keypress, here the keypress occurs every 10 minutes (600 Seconds).
    .EXAMPLE
       KeepAliveHours -Query
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 19.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 14.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 9.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 4.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around -0.04 Minutes
    
       KeepAlive has now completed... job will be cleaned up.
    
       KeepAlive has now completed.
    
       Run with the Query Switch to get an update on how long the timout will have to run.
    .EXAMPLE
       KeepAliveHours -Query
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 19.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 14.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 9.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around 4.96 Minutes
       Job will run till 09/30/2012 17:20:05 + 5 minutes, around -0.04 Minutes
    
       KeepAlive has now completed... job will be cleaned up.
    
       KeepAlive has now completed.
       
       The Query switch will also clean up the background job if you run this once the KeepAlive has complete..EXAMPLE
    .EXAMPLE
       KeepAliveHours -EndJob
       KeepAlive has now ended...
       
       Run Endjob once you download has complete to stop the Keepalive and remove the background job.
    .EXAMPLE
       KeepAliveHours -EndJob
       KeepAlive has now ended...
    
       Run EndJob anytime to stop the KeepAlive and remove the Job.
    .INPUTS
       KeepAliveHours - The time the keepalive will be active on the system
    .INPUTS
       SleepSeconds - The time between Keypresses. This should be less than the timeout of your computer screensaver or lock screen.
    .OUTPUTS
       This cmdlet creates a background job, when you Query the results the status from the background job will be outputed on the screen to let you know how long the KeepAlive will run for.
    .NOTES
       General notes
    .COMPONENT
       This is a standlone cmdlet, you may change the keystroke to do something more meaningful in a different scenario that this was originally written.
    .ROLE
       This utility should only be used in the privacy of your own home or locked office.
    .FUNCTIONALITY
       Call this function to enable a temporary KeepAlive for your computer. Allow you to download a large file without sleepin the computer.
    
       If the KeepAlive ends and you do not run -Query or -EndJob, then the completed job will remain.
    
       You can run Get-Job to view the job. Get-Job -Name KeepAlive | Remove-Job will cleanup the Job.
    
       By default you cannot create more than one KeepAlive Job, unless you provide a different JobName. There should be no reason to do this. With Query or EndJob, you can cleanup any old Jobs and then create a new one.
    .LINK
       https://gallery.technet.microsoft.com/scriptcenter/Keep-Alive-Simulates-a-key-9b05f980
    #>
        param (
            $KeepAliveHours = 1,
            $SleepSeconds = 180,
            $JobName = "KeepAlive",
            [Switch]$EndJob,
            [Switch]$Query,
            $KeyToPress = '{SCROLLLOCK}'   # Original function pressed  <Ctrl> = '^', changed this to ScrollLock
            # Reference for other keys: http://msdn.microsoft.com/en-us/library/office/aa202943(v=office.10).aspx
        )
    
        BEGIN {
            $Endtime = (Get-Date).AddHours($KeepAliveHours)
        } # BEGIN
    
        PROCESS {
            # Manually end the job and stop the KeepAlive.
            if ($EndJob) {
                if (Get-Job -Name $JobName -ErrorAction SilentlyContinue) {
                    Stop-Job -Name $JobName
                    Remove-Job -Name $JobName
                    "`n$JobName has now ended..."
                }
                else { "`nNo job $JobName." }
            }
            # Query the current status of the KeepAlive job.
            elseif ($Query) {
                try {
                    if ((Get-Job -Name $JobName -ErrorAction Stop).PSEndTime) {
                        Receive-Job -Name $JobName
                        Remove-Job -Name $JobName
                        "`n$JobName has now completed."
                    }
                    else { Receive-Job -Name $JobName -Keep }
                }
                catch {
                    Receive-Job -Name $JobName -ErrorAction SilentlyContinue
                    "`n$JobName has ended..."
                    Get-Job -Name $JobName -ErrorAction SilentlyContinue | Remove-Job
                }
            }
            # Start the KeepAlive job.
            elseif (Get-Job -Name $JobName -ErrorAction SilentlyContinue) {
                "`n$JobName already started, please use: Start-Keepalive -Query"
            }
            else {
                $Job = {
                    param ($Endtime,$SleepSeconds,$JobName,$KeyToPress)
                    "`nStarttime is $(Get-Date)"
    
                    While ((Get-Date) -le (Get-Date $EndTime)) {
                        # Wait SleepSeconds to press (This should be less than the screensaver timeout)
                        Start-Sleep -Seconds $SleepSeconds
                        $Remaining = [Math]::Round( ( (Get-Date $Endtime) - (Get-Date) | Select-Object -ExpandProperty TotalMinutes ),2 )
                        "Job will run till $EndTime + $([Math]::Round( $SleepSeconds/60, 2 )) minutes, around $Remaining Minutes"
                        # This is the sending of the KeyStroke
                        $x = New-Object -COM WScript.Shell
                        $x.SendKeys($KeyToPress)
                    }
                    try {
                        "`n$JobName has now completed... job will be cleaned up."
    
                        # Would be nice if the job could remove itself, below will not work.
                        # Receive-Job -AutoRemoveJob -Force
                        # Still working on a way to automatically remove the job
                    }
                    catch { "Something went wrong, manually remove job $JobName" }
                } # Job
    
                $JobProperties =@{   # Hash table
                    ScriptBlock  = $Job
                    Name         = $JobName
                    ArgumentList = $Endtime,$SleepSeconds,$JobName,$KeyToPress
                }
                Start-Job @JobProperties
                "`nKeepAlive set to run until $EndTime"
            }
        } # PROCESS
    } # Start-KeepAlive
    
    
# if ($computertype -eq Server) {
#     do not enable ProfileExtensions or load Custom-Tools.psm1
#     Load .psm1 from an ad hoc directory?
#     Create function in $Profile Get-Extensions / Toolkit to enable them
# }
# If server + administrator, then WARNING

# Windows Registry Editor Version 5.00
# 
# [HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server]
# "KeepAliveEnable"=dword:00000001

function Update-AllModules {
    # Update all installed Modules to latest versions with feedback
    # https://github.com/itpro-tips/PowerShell-Toolbox/blob/master/Update-AllPowerShellModules.ps1
    Write-Host -ForegroundColor cyan 'Define PowerShell to use TLS1.2 in this session, needed since 1st April 2020 (https://devblogs.microsoft.com/powershell/powershell-gallery-tls-support/)'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    # If required: Register PSGallery PSprovider and set as Trusted source:
    #    Register-PSRepository -Default -ErrorAction SilentlyContinue
    #    Set-PSRepository -Name PSGallery -InstallationPolicy trusted -ErrorAction SilentlyContinue

    $modules = Get-InstalledModule

    foreach ($module in $modules.Name) {
        $currentVersion = $null
    
        if ($null -ne (Get-InstalledModule -Name $module -ErrorAction SilentlyContinue)) {
            $currentVersion = (Get-InstalledModule -Name $module -AllVersions).Version
        }
    
        $moduleInfos = Find-Module -Name $module
    
        if ($null -eq $currentVersion) {
            Write-Host -ForegroundColor Cyan "Install from PowerShellGallery : $($moduleInfos.Name) - $($moduleInfos.Version). Release date: $($moduleInfos.PublishedDate)"  
        
            try {
                Install-Module -Name $module -Force
            }
            catch {
                Write-Host -ForegroundColor red "$_.Exception.Message"
            }
        }
        elseif ($moduleInfos.Version -eq $currentVersion) {
            Write-Host -ForegroundColor Green "$($moduleInfos.Name) already installed in the latest version ($currentVersion. Release date: $($moduleInfos.PublishedDate))"
        }
        elseif ($currentVersion.count -gt 1) {
            Write-Warning "$module is installed in $($currentVersion.count) versions (versions: $currentVersion)"
            Write-Host -ForegroundColor Cyan "Uninstall all $module PowerShell module versions"

            try {
                Get-InstalledModule -Name $module -AllVersions | Uninstall-Module -Force
            }
            catch {
                Write-Host -ForegroundColor red "$_.Exception.Message"
            }

            Write-Host -ForegroundColor Cyan "Install from PowerShellGallery : $($moduleInfos.Name) - $($moduleInfos.Version). Release date: $($moduleInfos.PublishedDate)"  
        
            try {
                Install-Module -Name $module -Force
            }
            catch {
                Write-Host -ForegroundColor red "$_.Exception.Message"
            }
        }
        else {       
            Write-Host -ForegroundColor Cyan "Update from PowerShellGallery from $currentVersion to $($moduleInfos.Name) - $($moduleInfos.Version). Release date: $($moduleInfos.PublishedDate)" 
            try {
                Update-Module -Name $module -Force
            }
            catch {
                Write-Host -ForegroundColor red "$_.Exception.Message"
            }
        }
    }
}



function Example-Splatting {
    # https://powershell.org/forums/topic/variables-in-write-host-command/
    "You can create a hash table with the names of the parameters and the values you wish to supply to"
    "those parameters then pass the hash table variable to the cmdlet prefixed with an @ symbol.`n"
    '$hash = @{ object = "test" ; foregroundcolor = "red" }'
    'Write-host @hash'
    $hash = @{ object = "test" ; foregroundcolor = "red" }
    Write-host @hash
    ""
    "Equivalent to:"
    $nonewline = $true
    $color = 'red'
    $count = " "*6
    write-host $count -BackgroundColor $color -NoNewline:$nonewline
}

function Convert-ImageToAsciiArt {
     param(
         [Parameter(Mandatory)] [String]$Path,
         [ValidateRange(20,20000)] [int]$MaxWidth=80,
         [float]$ratio = 1.5   # Character height:width ratio
    )

    # https://www.nextofwindows.com/turning-any-image-file-to-an-ascii-art-in-powershell
    # https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/create-ascii-art

    # load drawing functionality 
    Add-Type -AssemblyName System.Drawing 
    $characters = '$#H&@*+;:-,. '.ToCharArray()   # characters from dark to light  
    $c = $characters.count 
    $image = [Drawing.Image]::FromFile($path )   # load image and get image size 
    [int]$maxheight = $image.Height / ($image.Width / $maxwidth) / $ratio
    $bitmap = new-object Drawing.Bitmap($image ,$maxwidth,$maxheight)   # paint image on a bitmap with the desired size 
    [System.Text.StringBuilder]$sb = ""   # use a string builder to store the characters
    for ([int]$y=0; $y -lt $bitmap.Height; $y++) {  # take each pixel line...  
        for ([int]$x=0; $x -lt $bitmap.Width; $x++) {   # take each pixel column... 
            # examine pixel 
            $color = $bitmap.GetPixel($x, $y)
            $brightness = $color.GetBrightness() 
            # choose the character that best matches the pixel brightness
            [int]$offset = [Math]::Floor($brightness*$c) 
            $ch = $characters[$offset] 
            if (-not $ch) { $ch = $characters[-1] }  
            # add character to line 
            $null = $sb.Append($ch)
        }
        $null = $sb.AppendLine()   # add a new line 
    } 
    $image.Dispose()   # clean up and return string 
    $sb.ToString()
}
# To call the function, generate the ASCII file and display it.
# $Path = "C:\Program Files\Microsoft Office\root\CLIPART\PUB60COR\J0313970.JPG"   # Black and white mountain
# $Path = "C:\Program Files\Microsoft Office\root\CLIPART\PUB60COR\J0099165.JPG"   # Cat drawing
# $Path = "C:\Program Files\Microsoft Office\root\CLIPART\PUB60COR\J0202045.JPG"   # Child
# $Path = "C:\Windows\Web\Screen\img100.jpg"   # Windows Lock Screen
# $Path = "C:\Windows\Web\Screen\img101.jpg"   # Windows Lock Screen
# $Path = "C:\Windows\Web\Screen\img102.jpg"   # Windows Lock Screen
# $Path = "C:\Windows\Web\Screen\img103.jpg"   # Windows Lock Screen
# $Path = "C:\Windows\Web\Screen\img104.jpg"   # Windows Lock Screen
# $Path = "C:\Windows\Web\Screen\img105.jpg"   # Windows Lock Screen
# $OutPath = "$env:temp\asciiart.txt"
# Convert-ImageToAsciiArt -Path $Path -MaxWidth 150 | Set-Content -Path $OutPath -Encoding UTF8
# Invoke-Item -Path $OutPath

function Help-VSCodeKeyboardShortcutsForWindows {
    # Get the official Microsoft VS Code Keyboard Shortcuts PDF and open it with default PDF viewer
    $url = 'https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf'
    $desktop = [Environment]::GetFolderPath('Desktop')   # get desktop path if want to place it there
    $destination = "$env:TEMP\VSCodeKeyboardShortcutsForWindows.pdf"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12   # enable TLS1.2 for HTTPS connections
    Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing   # Download PDF file
    Invoke-Item -Path $destination   # Open downloaded file in associated program
}

##################################################
#  Linux top equivalent in PowerShell
##################################################
# https://superuser.com/questions/176624/linux-top-command-for-windows-powershell
# TopPS -SortCol CPU -top 10      # Show top 10 items sorted by CPU usage
# TopPS -SortCol Memory -top 10   # Show top 10 items sorted by Memory usage
# While (1) { $p = TopPS CPU 10 ; cls ; $p }
# Alternatives:   Note 'resmon' / 'taskman' to start the Resource Monitor / Task Manager (or Ctrl-Alt-Esc)
# while (1) { ps | findstr explorer | sort -desc cpu | select -first 30; sleep -seconds 2; cls }   # Filter by process
# while (1) { ps | Sort-Object -Property cpu -Descending | select -First 10 ; Write-Host "`nRefresh in 3 sec...`n" ; sleep -Seconds 3 }

# Like Top, this should stay in same position on screen
# $saveY = [console]::CursorTop ; # $saveX = [console]::CursorLeft      
# while ($true) { Get-Process | Sort -Descending CPU | Select -First 10 ; Sleep -Seconds 2 ; [console]::setcursorposition($saveX,$saveY+3) }

function TopPS ([string]$SortCol = "Memory", [int]$Top = 30) {

    if ($args[0] -eq $null) { $SortCol = "Memory" } else { $SortCol = $args[0] }
    if ($args[1] -eq $null) { $Top = 10 } else { $Top = $args[1] }

    $LogicalProcessors = (Get-WmiObject -class Win32_processor -Property NumberOfLogicalProcessors).NumberOfLogicalProcessors;

    ## Check user level of PowerShell 
    if ( ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) ) {
        $procTbl = get-process -IncludeUserName | select ID, Name, UserName, Description, MainWindowTitle
    } else {
        $procTbl = get-process | select ID, Name, Description, MainWindowTitle
    }

    Get-Counter `
        '\Process(*)\ID Process',`
        '\Process(*)\% Processor Time',`
        '\Process(*)\Working Set - Private'`
        -ea SilentlyContinue |
    foreach CounterSamples |
    where InstanceName -notin "_total","memory compression" |
    group { $_.Path.Split("\\")[3] } |
    foreach {
        $procIndex = [array]::indexof($procTbl.ID, [Int32]$_.Group[0].CookedValue)
        [pscustomobject]@{
            Name = $_.Group[0].InstanceName;
            ID = $_.Group[0].CookedValue;
            User = $procTbl.UserName[$procIndex]
            CPU = if($_.Group[0].InstanceName -eq "idle") {
                $_.Group[1].CookedValue / $LogicalProcessors 
                } else {
                $_.Group[1].CookedValue 
                };
            Memory = $_.Group[2].CookedValue / 1KB;
            Description = $procTbl.Description[$procIndex];
            Title = $procTbl.MainWindowTitle[$procIndex];
        }
    } |
    sort -des $SortCol |
    select -f $Top @(
        "Name", "ID", "User",
        @{ n = "CPU"; e = { ("{0:N1}%" -f $_.CPU) } },
        @{ n = "Memory"; e = { ("{0:N0} K" -f $_.Memory) } },
        "Description", "Title"
        ) | ft -a
}

function ErrBlack {
    # Change Error codes to White on Red background for better readability
    $Host.PrivateData.ErrorBackgroundColor = "Black"
    $Host.PrivateData.ErrorForegroundColor = "Red"

    # To make your color changes permanent, suggestion was to add the commands to one of your PowerShell profiles.
    # But I cannot make the below work reliably to have a black console with white text.
    # $Shell = $Host.UI.RawUI
    # $Shell.BackgroundColor = "Black"
    # $Shell.ForegroundColor = "White"
    # $Shell.CursorSize = 10
}

function Lock-Toolkit {
    # 1. Compress toolkit files (preferably with password though PS 5 tools do not allow that)
    # 2. Remove the Toolkit line from $profile
    # 3. Possibly remove scripts and modules
    # 4. Create a rebuild script in the Scripts folder to recreate the Toolkit (or just use iex)

    # PowerShell 5.0 (from "Microsoft.PowerShell.Archive")
    # Compress-Archive / Expand-Archive
    # Create result.zip from the entire Test folder:
    # Compress-Archive -Path C:\Test -DestinationPath C:\result
    # Extract the content of result.zip in the specified Test folder:
    # Expand-Archive -Path result.zip -DestinationPath C:\Test

    # Add-Type -A System.IO.Compression.FileSystem
    # [IO.Compression.ZipFile]::CreateFromDirectory('foo', 'foo.zip')
    # [IO.Compression.ZipFile]::ExtractToDirectory('foo.zip', 'bar')
    
    # If Java installed, compress to a zip using the jar command:
    # jar -cMf targetArchive.zip sourceDirectory
    # c = Creates a new archive file, M = do not add manifest file to archive, f = target file name.
    # Might not need the "-" in "-cMf"

    # Using 7-Zip:
    # Zip: you have a folder foo, and want to zip it to myzip.zip
    # "C:\Program Files\7-Zip\7z.exe" a  -r myzip.zip -w foo -mem=AES256
    # Unzip: you want to unzip it (myzip.zip) to current directory (./)
    # "C:\Program Files\7-Zip\7z.exe" x  myzip.zip  -o./ -y -r
}





# More Prompt options: https://mshforfun.blogspot.com/2006/05/perfect-prompt-for-windows-powershell.html
# https://devblogs.microsoft.com/scripting/weekend-scripter-customize-powershell-title-and-prompt/
# https://github.com/AmrEldib/cmder-powershell-powerline-prompt
# https://hodgkins.io/ultimate-powershell-prompt-and-git-setup
# https://gallery.technet.microsoft.com/scriptcenter/Custom-PowerShell-GUI-7c7fbda8
# https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/setting-powershell-title-text
# Very over the top console prompt project
# https://github.com/myleftshoe/powershell/blob/Powerline-style/.profile.ps1
# https://dev.to/hf-solutions/how-to-uniquify-your-powershell-profile-2b35
# Customise Console
# https://devblogs.microsoft.com/scripting/customize-the-powershell-console-for-increased-efficiency/
# <https://stackoverflow.com/questions/15694338/how-to-get-a-list-of-custom-powershell-functions>
# PS Core 6 / 7 propmpt:
# https://jdhitsolutions.com/blog/powershell/6388/maximizing-my-prompt-in-powershell-core/
# https://stackoverflow.com/questions/5725888/windows-powershell-changing-the-command-prompt
# https://github.com/dahlbyk/posh-git/wiki/Customizing-Your-PowerShell-Prompt
# https://www.norlunn.net/2019/10/07/powershell-customize-the-prompt/
# https://devblogs.microsoft.com/scripting/customize-the-powershell-console-for-increased-efficiency/
# https://www.itprotoday.com/powershell/powershell-basics-console-configuration
# https://www.jonathanmedd.net/2015/01/how-to-make-use-of-functions-in-powershell.html
# https://jdhitsolutions.com/blog/powershell/6388/maximizing-my-prompt-in-powershell-core/

function PromptXXX {
    "$($MyInvocation.InvocationName)"
    "$(($MyInvocation.MyCommand).Name)"
    ""
    if ($MyInvocation.InvocationName -eq ".") {
        "dotsourced"
    } else {
        "`nWarning: Must dotsource '$($MyInvocation.MyCommand)' or it will not be applied to this session.`n`n   . $($MyInvocation.MyCommand)`n"
    }
}

# if ($MyInvocation.InvocationName -eq ".") {
#     "Dotsourced!"
# } else {
#     "`nWarning: Must dotsource '$($MyInvocation.MyCommand)' or it will not be applied to this session.`n`n   . $($MyInvocation.MyCommand)`n"
# }

function PromptDefault {
    # get-help about_Prompt
    # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_prompts?view=powershell-7
    function global:prompt {
        "PS $($executionContext.SessionState.Path.CurrentLocation)$('>' * ($nestedPromptLevel + 1)) ";
        # .Link
        # https://go.microsoft.com/fwlink/?LinkID=225750
        # .ExternalHelp System.Management.Automation.dll-help.xml

        $Elevated = ""
        $user = [Security.Principal.WindowsIdentity]::GetCurrent();
        if ((New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {$Elevated = "Administrator: "}
        # $TitleVer = "PS v$($PSVersionTable.PSversion.major).$($PSVersionTable.PSversion.minor)"
        $TitleVer = "PowerShell"
        $Host.UI.RawUI.WindowTitle = "$($Elevated)$($TitleVer)"
    }
}

# More simple alternative prompt, need to dotsource this
function PromptTimeUptime {
    function global:prompt {
        # Adds date/time to prompt and uptime to title bar
        $Elevated = "" ; if (Test-Admin) {$Elevated = "Administrator: "}
        $up = Uptime
        $Host.UI.RawUI.WindowTitle = $Elevated + "PowerShell [Uptime: $up]"   # Title bar info
        $path = Get-Location
        Write-Host '[' -NoNewline
        Write-Host (Get-Date -UFormat '%T') -ForegroundColor Green -NoNewline   # $TitleDate = Get-Date -format "dd/MM/yyyy HH:mm:ss"
        Write-Host '] ' -NoNewline
        Write-Host "$path" -NoNewline
        return "> "   # Must have a line like this at end of prompt or you always get " PS>" on the prompt
    }
}

function PromptTruncatedPaths {
    # https://www.johndcook.com/blog/2008/05/12/customizing-the-powershell-command-prompt/
    function global:prompt {
        $cwd = (get-location).Path
        [array]$cwdt=$()
        $cwdi = -1
        do {$cwdi = $cwd.indexofany("\", $cwdi+1) ; [array]$cwdt+=$cwdi} until($cwdi -eq -1)
        if ($cwdt.count -gt 3) { $cwd = $cwd.substring(0,$cwdt[0]) + ".." + $cwd.substring($cwdt[$cwdt.count-3]) }
        $host.UI.RawUI.WindowTitle = "$(hostname) - $env:USERDNSDOMAIN$($env:username)"
        $host.UI.Write("Yellow", $host.UI.RawUI.BackGroundColor, "[PS]")
        " $cwd> "
    }
}

function PromptShortenPath {
    # https://stackoverflow.com/questions/1338453/custom-powershell-prompts
    function global:shorten-path([string] $path) {
        $loc = $path.Replace($HOME, '~')
        # remove prefix for UNC paths
        $loc = $loc -replace '^[^:]+::', ''
        # make path shorter like tabs in Vim,
        # handle paths starting with \\ and . correctly
        return ($loc -replace '\\(\.?)([^\\])[^\\]*(?=\\)','\$1$2')
    }
    function global:prompt {
        # our theme
        $cdelim = [ConsoleColor]::DarkCyan
        $chost = [ConsoleColor]::Green
        $cloc = [ConsoleColor]::Cyan

        write-host "$([char]0x0A7) " -n -f $cloc
        write-host ([net.dns]::GetHostName()) -n -f $chost
        write-host ' {' -n -f $cdelim
        write-host (shorten-path (pwd).Path) -n -f $cloc
        write-host '}' -n -f $cdelim
        return ' '
    }
}

function PromptUserAndExecutionTimer {
    function global:prompt {

        ### Title bar info
        $user = [Security.Principal.WindowsIdentity]::GetCurrent();
        $Elevated = ""
        if ((New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {$Elevated = "Admin: "}
        $TitleVer = "PS v$($PSVersionTable.PSversion.major).$($PSVersionTable.PSversion.minor)"
        # $($executionContext.SessionState.Path.CurrentLocation.path)

        ### Custom Uptime without seconds (not really necessary)
        # $wmi = gwmi -class Win32_OperatingSystem -computer "."
        # $LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
        # [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
        # $s = "" ; if ($uptime.Days -ne 1) {$s = "s"}
        # $TitleUp = "[Up: $($uptime.days) day$s $($uptime.hours) hr $($uptime.minutes) min]"

        $Host.UI.RawUI.WindowTitle = "$($Elevated) $($TitleVer)"   # $($TitleUp)"

        ### History ID
        $HistoryId = $MyInvocation.HistoryId
        # Uncomment below for leading zeros
        # $HistoryId = '{0:d4}' -f $MyInvocation.HistoryId
        Write-Host -Object "$HistoryId " -NoNewline -ForegroundColor Cyan

    
        ### Time calculation
        $Success = $?
        $LastExecutionTimeSpan = if (@(Get-History).Count -gt 0) {
            Get-History | Select-Object -Last 1 | ForEach-Object {
                New-TimeSpan -Start $_.StartExecutionTime -End $_.EndExecutionTime
            }
        }
        else {
            New-TimeSpan
        }
    
        $LastExecutionShortTime = if ($LastExecutionTimeSpan.Days -gt 0) {
            "$($LastExecutionTimeSpan.Days + [Math]::Round($LastExecutionTimeSpan.Hours / 24, 2)) d"
        }
        elseif ($LastExecutionTimeSpan.Hours -gt 0) {
            "$($LastExecutionTimeSpan.Hours + [Math]::Round($LastExecutionTimeSpan.Minutes / 60, 2)) h"
        }
        elseif ($LastExecutionTimeSpan.Minutes -gt 0) {
            "$($LastExecutionTimeSpan.Minutes + [Math]::Round($LastExecutionTimeSpan.Seconds / 60, 2)) m"
        }
        elseif ($LastExecutionTimeSpan.Seconds -gt 0) {
            "$($LastExecutionTimeSpan.Seconds + [Math]::Round($LastExecutionTimeSpan.Milliseconds / 1000, 1)) s"
        }
        elseif ($LastExecutionTimeSpan.Milliseconds -gt 0) {
            "$([Math]::Round($LastExecutionTimeSpan.TotalMilliseconds, 0)) ms"
            # ms are 1/1000 of a sec so no point in extra decimal places here
        }
        else {
            "0 s"
        }
    
        if ($Success) {
            Write-Host -Object "[$LastExecutionShortTime] " -NoNewline -ForegroundColor Green
        }
        else {
            Write-Host -Object "! [$LastExecutionShortTime] " -NoNewline -ForegroundColor Red
        }
    
        ### User, removed
        $IsAdmin = (New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
        # Write-Host -Object "$($env:USERNAME)$(if ($IsAdmin){ '[A]' } else { '[U]' }) " -NoNewline -ForegroundColor DarkGreen
        # Write-Host -Object "$($env:USERNAME)" -NoNewline -ForegroundColor DarkGreen
        # Write-Host -Object " [" -NoNewline
        # if ($IsAdmin) { Write-Host -Object 'A' -NoNewline -F Red } else { Write-Host -Object 'U' -NoNewline }
        # Write-Host -Object "] " -NoNewline
        Write-Host "$($env:USERNAME)" -NoNewline -ForegroundColor DarkGreen
        Write-Host "[" -NoNewline
        if ($IsAdmin) { Write-Host 'A' -NoNewline -F Red } else { Write-Host -Object 'U' -NoNewline }
        Write-Host "] " -NoNewline
    
        # ### Path
        # $Drive = $pwd.Drive.Name
        # $Pwds = $pwd -split "\\" | Where-Object { -Not [String]::IsNullOrEmpty($_) }
        # $PwdPath = if ($Pwds.Count -gt 3) {
        #     $ParentFolder = Split-Path -Path (Split-Path -Path $pwd -Parent) -Leaf
        #     $CurrentFolder = Split-Path -Path $pwd -Leaf
        #     "..\$ParentFolder\$CurrentFolder"
        # go  # }
        # elseif ($Pwds.Count -eq 3) {
        #     $ParentFolder = Split-Path -Path (Split-Path -Path $pwd -Parent) -Leaf
        #     $CurrentFolder = Split-Path -Path $pwd -Leaf
        #     "$ParentFolder\$CurrentFolder"
        # }
        # elseif ($Pwds.Count -eq 2) {
        #     Split-Path -Path $pwd -Leaf
        # }
        # else { "" }
        # Write-Host -Object "$Drive`:\$PwdPath" -NoNewline

        Write-Host $pwd -NoNewline
        return "> "
    }
}

function PromptSlightlyBroken {
    # https://community.spiceworks.com/topic/1965997-custom-cmd-powershell-prompt

    # if ($MyInvocation.InvocationName -eq "PromptOverTheTop") {
    #     "`nWarning: Must dotsource '$($MyInvocation.MyCommand)' or it will not be applied to this session.`n`n   . $($MyInvocation.MyCommand)`n"
    # } else {
    if ($host.name -eq 'ConsoleHost') {
        # fff
        $Shell = $Host.UI.RawUI
        $Shell.BackgroundColor = "Black"
        $Shell.ForegroundColor = "White"
        $Shell.CursorSize = 10
    }
    # $Shell=$Host.UI.RawUI
    # $size=$Shell.BufferSize
    # $size.width=120
    # $size.height=3000
    # $Shell.BufferSize=$size
    # $size=$Shell.WindowSize
    # $size.width=120
    # $size.height=30
    # $Shell.WindowSize=$size
    # $Shell.BackgroundColor="Black"
    # $Shell.ForegroundColor="White"
    # $Shell.CursorSize=10
    # $Shell.WindowTitle="Console PowerShell"

    function global:Get-Uptime {

        # $wmi = gwmi -class Win32_OperatingSystem -computer "."   # Removed this method as not CIM compliant
        # $LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
        # [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
        $BootUpTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
        $CurrentDate = Get-Date
        $Uptime = $CurrentDate - $BootUpTime
        $s = "" ; if ($Uptime.Days -ne 1) {$s = "s"}
        $uptime_string = "$($uptime.days) day$s $($uptime.hours) hr $($uptime.minutes) min $($uptime.seconds) sec"
        # $os = Get-WmiObject win32_operatingsystem
        # $uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))
        # $days = $Uptime.Days ; if ($days -eq "1") { $days = "$days day" } else { $days = "$days days"}
        # $hours = $Uptime.Hours ; if ($hours -eq "1") { $hours = "$hours hr" } else { $hours = "$hours hrs"}
        # $minutes = $Uptime.Minutes ; if ($minutes -eq "1") { $minutes = "$minutes min" } else { $minutes = "$minutes mins"}
        # $Display = "$days, $hours, $minutes"
        Write-Output $uptime_string
    }
    function Spaces ($numspaces) { for ($i = 0; $i -lt $numspaces; $i++) { Write-Host " " -NoNewline } }

    # $MaximumHistoryCount=1024
    $IPAddress = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].IPAddress[0]
    $IPGateway = @(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].DefaultIPGateway[0]
    $UserDetails = "$env:UserDomain\$env:UserName (PS-HOME: $HOME)"
    $PSExecPolicy = Get-ExecutionPolicy
    $PSVersion = "$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor) ($PSExecPolicy)"
    $ComputerAndLogon = "$($env:COMPUTERNAME)"
    $ComputerAndLogonSpaces = 28 - $ComputerAndLogon.Length
    Clear
    Write-Host "-----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Green
    Write-Host "|    ComputerName:  " -nonewline -ForegroundColor Green; Write-Host $ComputerAndLogon -nonewline -ForegroundColor White ; Spaces $ComputerAndLogonSpaces ; Write-Host "UserName:" -nonewline -ForegroundColor Green ; Write-Host "  $UserDetails" -ForegroundColor White
    Write-Host "|    Logon Server:  " -nonewline -ForegroundColor Green; Write-Host $($env:LOGONSERVER)"`t`t`t`t" -nonewline -ForegroundColor White ; Write-Host "IP Address:`t" -nonewline -ForegroundColor Green ; Write-Host "`t$IPAddress ($IPGateway)" -ForegroundColor White
    Write-Host "|    Uptime:        " -nonewline -ForegroundColor Green; Write-Host "$(Get-Uptime)`t" -nonewline -ForegroundColor White; Write-Host "PS Version:`t" -nonewline -ForegroundColor Green ; Write-Host "`t$PSVersion" -ForegroundColor White
    Write-Host "-----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Green
    # Write-Host "-----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Green
    # Write-Host "|`tComputerName:`t" -nonewline -ForegroundColor Green; Write-Host $($env:COMPUTERNAME)"`t`t`t`t" -nonewline -ForegroundColor White ; Write-Host "UserName:`t$UserDetails" -ForegroundColor White
    # Write-Host "|`tLogon Server:`t" -nonewline -ForegroundColor Green; Write-Host $($env:LOGONSERVER)"`t`t`t`t" -nonewline -ForegroundColor White ; Write-Host "IP Address:`t$IPAddress ($IPGateway)" -ForegroundColor White
    # Write-Host "|`tUptime:`t`t" -nonewline -ForegroundColor Green; Write-Host "$(Get-Uptime)`t" -nonewline -ForegroundColor White; Write-Host "PS Version:`t$PSVersion" -ForegroundColor White
    # Write-Host "-----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Green
    function global:admin {
        $Elevated = ""
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
        if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $true) { $Elevated = "Administrator: " }
        $Host.UI.RawUI.WindowTitle = "$Elevated$TitleVer"
    }
    admin
    Set-Location C:\

    function global:prompt{
        $br = "`n"
        Write-Host "[" -noNewLine
        Write-Host $(Get-date) -ForegroundColor Green -noNewLine
        Write-Host "] " -noNewLine
        Write-Host "[" -noNewLine
        Write-Host "$env:username" -Foregroundcolor Red -noNewLine
        Write-Host "] " -noNewLine
        Write-Host "[" -noNewLine
        Write-Host $($(Get-Location).Path.replace($home,"~")) -ForegroundColor Yellow -noNewLine
        Write-Host $(if ($nestedpromptlevel -ge 1) { '>>' }) -noNewLine
        Write-Host "] "
        return "> "
    }
}

Set-Alias p0 PromptDefault
Set-Alias p-default PromptDefault
Set-Alias p-timer PromptUserAndExecutionTimer   # Using this as my console default
Set-Alias p-shortpath PromptShortenPath
Set-Alias p-truncpath PromptTruncatedPaths
Set-Alias p-uptime PromptTimeUptime
Set-Alias p-ott PromptSlightlyBroken

# View current prompt with: (get-item function:prompt).scriptblock   or   cat function:\prompt


# function PromptOverTheTop {
#     # https://community.spiceworks.com/topic/1965997-custom-cmd-powershell-prompt
#     function global:admin {
#         $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
#         if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $true) {
#             $Host.UI.RawUI.WindowTitle = "Administrator: " + $host.UI.RawUI.WindowTitle
#         }
#     }
#     function global:Get-Uptime {
#         $os = Get-WmiObject win32_operatingsystem
#         $uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))
#         $days = $Uptime.Days ; if ($days -eq "1") { $days = "$days day" } else { $days = "$days days"}
#         $hours = $Uptime.Hours ; if ($hours -eq "1") { $hours = "$hours hr" } else { $hours = "$hours hrs"}
#         $minutes = $Uptime.Minutes ; if ($minutes -eq "1") { $minutes = "$minutes min" } else { $minutes = "$minutes mins"}
#         $Display = "$days, $hours, $minutes"
#         Write-Output $Display
#     }
# 
#     if ($host.name -eq 'ConsoleHost') {
#         fff
#         $Shell = $Host.UI.RawUI
#         $Shell.BackgroundColor = "Black"
#         $Shell.ForegroundColor = "White"
#         $Shell.CursorSize = 10
#     }
#     clear
#     # $MaximumHistoryCount=1024
#     $IPAddress=@(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].IPAddress[0]
#     $IPGateway=@(Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.DefaultIpGateway})[0].DefaultIPGateway[0]
#     $PSExecPolicy=Get-ExecutionPolicy
#     $PSVersion=$PSVersionTable.PSVersion.Major
#     Write-Host "-----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Green
#     Write-Host "|`tComputerName:`t`t" -nonewline -ForegroundColor Green ; Write-Host $($env:COMPUTERNAME)"`t`t`t`t" -nonewline -ForegroundColor white ; Write-Host "UserName:`t`t" -nonewline -ForegroundColor Green       ; Write-Host $env:UserDomain\$env:UserName"`t`t" -ForegroundColor white
#     Write-Host "|`tLogon Server:`t`t" -nonewline -ForegroundColor Green ; Write-Host $($env:LOGONSERVER)"`t`t`t`t" -nonewline -ForegroundColor white  ; Write-Host "IP Address:`t`t" -nonewline -ForegroundColor Green     ; Write-Host $IPAddress"`t`t" -ForegroundColor white
#     Write-Host "|`tUptime:`t`t`t" -nonewline -ForegroundColor Green     ; Write-Host $(Get-Uptime)"`t" -nonewline -ForegroundColor white              ; Write-Host "`tPS Version:`t`t`t" -nonewline -ForegroundColor Green ; Write-Host $PSVersion"" -ForegroundColor white
#     Write-Host "-----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Green
#     $br
#     Set-Location C:\
#     function global:prompt{
#         if ($host.name -eq 'ConsoleHost') {
#             $Shell = $Host.UI.RawUI
#             $Shell.BackgroundColor = "Black"
#             $Shell.ForegroundColor = "White"
#             $Shell.CursorSize = 10
#         }
#         admin
#         $br = "`n"
#         Write-Host "[" -noNewLine
#         Write-Host $(Get-date) -ForegroundColor Green -noNewLine
#         Write-Host "] " -noNewLine
#         Write-Host "[" -noNewLine
#         Write-Host "$env:username" -Foregroundcolor Red -noNewLine
#         Write-Host "] " -noNewLine
#         Write-Host "[" -noNewLine
#         Write-Host $($(Get-Location).Path.replace($home,"~")) -ForegroundColor Yellow -noNewLine
#         Write-Host $(if ($nestedpromptlevel -ge 1) { '>>' }) -noNewLine
#         Write-Host "] " 
#         $Shell = $Host.UI.RawUI
#         $Shell.BackgroundColor = "Black"
#         $Shell.ForegroundColor = "White"
#         $Shell.CursorSize = 10
#         return "> "
#     }
# }
# # $Shell=$Host.UI.RawUI
# # $size=$Shell.BufferSize
# # $size.width=120
# # $size.height=3000
# # $Shell.BufferSize=$size
# # $size=$Shell.WindowSize
# # $size.width=120
# # $size.height=30
# # $Shell.WindowSize=$size
# # $Shell.BackgroundColor="Black"
# # $Shell.ForegroundColor="White"
# # $Shell.CursorSize=10
# # $Shell.WindowTitle="Console PowerShell"



###   # Customisation with Posh-Git ...
###   # $GitPromptSettings.DefaultPromptWriteStatusFirst = $true
###   # $GitPromptSettings.DefaultPromptBeforeSuffix.Text = '`n$([DateTime]::now.ToString("MM-dd HH:mm:ss"))'
###   # $GitPromptSettings.DefaultPromptBeforeSuffix.ForegroundColor = 0x808080
###   # $GitPromptSettings.DefaultPromptSuffix = ' $((Get-History -Count 1).id + 1)$(">" * ($nestedPromptLevel + 1)) '
###   
###   # More advanced:
###   # https://hodgkins.io/ultimate-powershell-prompt-and-git-setup
###   
###   # Alternatives:
###   # $host.UI.RawUI.WindowTitle = "Windows PowerShell: ...  $path  ..."
###   # Get current time: # $date = Get-Date -Format 'ddd, MMM dd' # $time = Get-Date -Format 'hh:mm:ss'
###   
###   # Changes PowerShell prompt to current folder only (shorter prompt) in lowercase.
###   # function prompt {
###   #     "$((Get-Location | Split-Path -leaf).ToLower())> "
###   # }
###   
###   # function global:prompt {    
###   #     # ### Path
###   #     # $Drive = $pwd.Drive.Name
###   #     # $Pwds = $pwd -split "\\" | Where-Object { -Not [String]::IsNullOrEmpty($_) }
###   #     # $PwdPath = if ($Pwds.Count -gt 3) {
###   #     #     $ParentFolder = Split-Path -Path (Split-Path -Path $pwd -Parent) -Leaf
###   #     #     $CurrentFolder = Split-Path -Path $pwd -Leaf
###   #     #     "..\$ParentFolder\$CurrentFolder"
###   #     # go  # }
###   #     # elseif ($Pwds.Count -eq 3) {
###   #     #     $ParentFolder = Split-Path -Path (Split-Path -Path $pwd -Parent) -Leaf
###   #     #     $CurrentFolder = Split-Path -Path $pwd -Leaf
###   #     #     "$ParentFolder\$CurrentFolder"
###   #     # }
###   #     # elseif ($Pwds.Count -eq 2) {
###   #     #     Split-Path -Path $pwd -Leaf
###   #     # }
###   #     # else { "" }
###   #     # Write-Host -Object "$Drive`:\$PwdPath" -NoNewline
###   # 
###   #     Write-Host $pwd -NoNewline
###   #     return "> "
###   # }
###   
###   #  View current prompt with: (get-item function:prompt).scriptblock   or   cat function:\prompt





# https://www.gngrninja.com/script-ninja/2016/3/20/powershell
# https://www.gngrninja.com/script-ninja/2019/11/7/visual-studio-code-powershell-setup
# http://akuederle.com/modify-your-powershell-prompt

# Customise Console
# https://stackoverflow.com/questions/16280402/setting-powershell-colors-with-hex-values-in-profile-script






# https://ridicurious.com/2016/10/19/powershell-customize-directory-path-in-ps-prompt/
# https://geekeefy.wordpress.com/2016/10/19/powershell-customize-directory-path-in-ps-prompt/
# https://www.gngrninja.com/script-ninja/2016/6/18/powershell-getting-started-part-12-creating-custom-objects
# PowerTab : https://codeyarns.com/2011/07/30/how-to-install-powertab-for-powershell/
# https://mcpmag.com/articles/2016/06/09/display-gui-message-boxes-in-powershell.aspx
# https://blogs.technet.microsoft.com/jamesone/2009/06/24/how-to-get-user-input-more-nicely-in-powershell/
# https://social.technet.microsoft.com/wiki/contents/articles/24030.powershell-demo-prompt-for-choice.aspx
# https://ilovepowershell.com/2010/07/23/profile-trick-how-to-set-administrator-mode-background-color/


#
# Learning Git stuff ...
# https://gist.github.com/jivkok/c62746ccdfc0695218a2
# git clone https://github.com/jivkok/dotfiles.git $HOME\dotfiles






# function PromptShortenPath {
#     # https://stackoverflow.com/questions/1338453/custom-powershell-prompts
#     function shorten-path([string] $path) {
#         $loc = $path.Replace($HOME, '~')
#         # remove prefix for UNC paths
#         $loc = $loc -replace '^[^:]+::', ''
#         # make path shorter like tabs in Vim,
#         # handle paths starting with \\ and . correctly
#         return ($loc -replace '\\(\.?)([^\\])[^\\]*(?=\\)','\$1$2')
#     }
#     function prompt {
#         # our theme
#         $cdelim = [ConsoleColor]::DarkCyan
#         $chost = [ConsoleColor]::Green
#         $cloc = [ConsoleColor]::Cyan
#         
#         write-host "$([char]0x0A7) " -n -f $cloc
#         write-host ([net.dns]::GetHostName()) -n -f $chost
#         write-host ' {' -n -f $cdelim
#         write-host (shorten-path (pwd).Path) -n -f $cloc
#         write-host '}' -n -f $cdelim
#         return ' '
#     }
#     if ($MyInvocation.InvocationName -eq "PromptShortenPath") {
#         "`nWarning: Must dotsource '$($MyInvocation.MyCommand)' or it will not be applied to this session.`n`n   . $($MyInvocation.MyCommand)`n"
#     } else {
#         . prompt 
#     }
# }

# Function Elevate-Process {
# Set-Alias -Name elevate -Value Elevate-Process
# Set-Alias -Name sudo -Value Elevate-Process
# Elevate-Process fails if $args is a Cmdlet. e.g.   sudo dir
# Exception calling "Start" with "1" argument(s): "The system cannot find the file specified"
# At D:\Gist\ProfileExtensions.ps1:700 char:5
# +     [System.Diagnostics.Process]::Start($Process);
# +     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#     + CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
#     + FullyQualifiedErrorId : Win32Exception
#
# Need to do a test, take $arg[0]. If that is a cmdlet, then *add* powershell.exe to the command?
# I think this might be the only way to deal with internal commands
# This is the other method often seen, I've not had as good results with this.
# but maybe I need a combination, use the below for PowerShell Cmdlets and above for other?
# # $MyInvocation.MyCommand.Path   #  + $MyInvocation.UnboundArguments
# if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
#     if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
#         echo "-File `"" + $Args + "`" "
#         $CommandLine = "-File `"" + $Args + "`""
#         Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
#     }
# }

# Collect various Elevate techniques here, I've saved the a very efficient one here in profile
# Note also the Modules: "Sudo", "PSSudo"
# https://gist.github.com/TaoK/1582185
# Simple version, http://weestro.blogspot.com/2009/08/sudo-for-powershell.html
# Need to find original source for this version which works well for me so far.
function Invoke-Elevate {
    <#
    TODO:
    - have output return to the main screen
    - launch the elevated process with wscript to avoid UAC
    - work out an elevated prompt, and all commands ran will use elevation until...
    #>

    [CmdLetBinding()]
    [CmdletBinding(DefaultParameterSetName='Command')]
    param
    (
        # ScriptBlock: Negates the need for Command
        [Parameter(Mandatory=$false,ParameterSetName="Command")]
        [Parameter(Mandatory=$true, Position=0,ParameterSetName='ScriptBlock',
        HelpMessage='Scriptblock of commands to be executed')]
        [ScriptBlock] $ScriptBlock,

        # Command: Negates the need for ScriptBlock
        [Parameter(Mandatory=$false, ParameterSetName='ScriptBlock')]
        [Parameter(Mandatory=$true, Position=0, ParameterSetName='Command',
        HelpMessage='Commands to be executed')]
        [String] $Command,

        [Switch] $NoProfile,

        [Switch] $Persist
    )

    begin
    {
        # Invoke-VariableBaseLine

        [Bool] $boolDebug = $PSBoundParameters.Debug.IsPresent
    }

    process
    {

        [String] $strCommand = "& { $ScriptBlock }"

        if ($Command)
        {
            [String] $strCommand = $Command
        }

        [String] $strEncodedCommand = [Convert]::ToBase64String($([System.Text.Encoding]::Unicode.GetBytes($strCommand)))
        [String] $strArguments = "-Exec ByPass -EncodedCommand $strEncodedCommand"

        if ($NoProfile)
        {
            $strArguments =+ ' -Nop'
        }

        if ($Persist)
        {
            $strArguments += ' -NoExit'
        }

        Start-Process PowerShell -Verb runas -ArgumentList $strArguments
    }

    end
    {
        # Invoke-VariableBaseLine -Clean
    }
}
Set-Alias sudo Invoke-Elevate
Set-Alias elevate Invoke-Elevate
function sudops { sudo powershell.exe }

# No parameters but will use $args[0], $args[1]
function New-BackgroundProcess {
    # Create the .NET objects
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $newproc = New-Object System.Diagnostics.Process
    # Basic stuff, process name and arguments
    $psi.FileName = $args[0]
    $psi.Arguments = $args[1]
    # Hide any window it might try to create
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = 'Hidden'
    # Set up and start the process
    $newproc.StartInfo = $psi
    $newproc.Start()
    # Return the process object to the caller
    $newproc
}


# For cd, note the PSProviders and how we can 'cd' into them:
# Get-PSProvider   # Defaults are: Registry, Alias, Environment, FileSystem, Function, Variable
# New-PSDrive -name G -PSProvider Registry -Root HKCU:\Software   # Create G:\ drive (psprovider) that points to 'Current User\Software' in the registry
# To see more functionaliy for PSProviders, get-help about_providers.
# Explicitly overriding 'cd' which is an alias of 'Set-Location'
# See "get-help about_Command_Precedence" for more details on the precedence rules.
# Should probably update this so that it saves the last location then a switch -back / -b to jump back in case of a mistake
# add-content ... some way to control the contents, or possibly read all, put into array, get the new one then write all back?

# Initially used a hashtable that was exposed in profile_extensions, but below system with CLIXML is better
# $gohash = @{
#     share   = $env:HOMESHARE
#     home    = "$HomeFix::$env:USERPROFILE"
#     homec   = $HomeFix
#     # inghome = "$(Split-Path (Split-Path $ProfileNetShare))\Desktop"
#     inghome = "\\ad.ing.net\WPS\NL\P\UD\200024\$HomeLeaf\Home"   # ING only!
# }

# gohash options:
# key-value hash table: could have multi-keys / mulit-values. This is possible but maybe too much hassle.
# 
# co same as cc but open the folder in windows explorer (could also add work locations etc)
# we can completely replace cd, if a shortcut is issued, 1. check if there is a *real* folder of that name in pwd, 2. use the gohash shortcut.
# Good to replace the cd shortcut in this way
# REGISTRY! can have shortcuts to various places in there!
# after a path is found, update the hash table to order so that the good path is first!
Function Convert-HashToString ($hashtable) { ($hashtable.GetEnumerator() | % { "$($_.Key)=$($_.Value)" }) -join ', ' }
Function Convert-ArrayToString ($array) { ($array.GetEnumerator() | % { "$($_)"}) -join ', ' }

function e ($file) {
    # Intelligent Edit, open file types according to criteria. e.g. .txt/.diz/.nfo open in Notepad++, but code files with code.exe
    # $ScriptFile = gci $MyInvocation.MyCommand.Path
    # $ScriptFull = $ScriptFile.FullName
    # $ScriptPath = Split-Path $ScriptFull -Parent
    # $ScriptName = Split-Path $ScriptFull -Leaf
    # $ScriptExt  = $ScriptFile.Extension
    # $ScriptNoExt = ($ScriptName -Split $ScriptExt)[0]       # Always use -Split for a string; .split('xyz') will split on that array of characters
    $FileExt = (gci $File).Extension
    if ($FileExt -in ".txt", ".nfo", ".diz") {
        if (Test-Path "C:\Program Files\Notepad++\notepad++.exe") { Start-Process "C:\Program Files\Notepad++\notepad++.exe" $File }
        else { Start-Process "C:\WINDOWS\notepad.exe" $File }
    }
    if ($FileExt -in ".ps1", ".psm1", ".psd1", ".py") {
        # Try default user-space VS Code first, then Program Files location, then Notepad++, and finally Notepad.
        if (Test-Path "$Home\AppData\Local\Programs\Microsoft VS Code\Code.exe") { Start-Process "$Home\AppData\Local\Programs\Microsoft VS Code\Code.exe" $File }
        elseif (Test-Path "C:\Program Files\Microsoft VS Code\Code.exe") { Start-Process "C:\Program Files\Microsoft VS Code\Code.exe" $File }
        elseif (Test-Path "C:\Program Files\Notepad++\notepad++.exe") { Start-Process "C:\Program Files\Notepad++\notepad++.exe" $File }
        else { Start-Process "C:\WINDOWS\notepad.exe" $File }
    }
}
function catw ($file) { Write-Wrap "$(cat $file)" }   # This needs work, carriage returns need to be respected or maybe a rule like: only respect carriage return if a full stop, space, then capital letter is involved which would handle most written text well.
function hosts { if (Test-Path "C:\Program Files\Notepad++\notepad++.exe") { Start-Process "C:\Program Files\Notepad++\notepad++.exe" "C:\Windows\System32\drivers\etc\hosts" } else { Start-Process "C:\WINDOWS\notepad.exe" "C:\Windows\System32\drivers\etc\hosts" } }
function profile { e $profile ; }
function edit ($file) { if (Test-Path "C:\Program Files\Notepad++\notepad++.exe") { Start-Process "C:\Program Files\Notepad++\notepad++.exe" $file } else { Start-Process "C:\WINDOWS\notepad.exe" $file } }
function mirror ($src, $dst) { robocopy $src $dst /mir /r:1 /w:1 /xjd /xf ntuser.dat*
    # /XJD :: eXclude Junction points and symbolic links for Directories.   <<< Quite important for home folder
    # /XJ :: eXclude Junction points and symbolic links, # /XJF :: eXclude symbolic links for Files,
    # /FFT :: assume FAT File Times (2-second granularity), # /DST :: compensate for one-hour DST time differences.
}

# Customised function that replicates the functionality of "Out-String -Stream" on the pipeline
# https://stackoverflow.com/questions/64630362/powershell-pipeline-compatible-select-string-function
function oss {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [psobject]$InputObject
    )
    begin
    {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }
            # set -Stream switch always
            $PSBoundParameters["Stream"] = $true
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Microsoft.PowerShell.Utility\Out-String', [System.Management.Automation.CommandTypes]::Cmdlet)
            $scriptCmd = {& $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        } catch {
            throw
        }
    }
    process
    {
        try {
            $steppablePipeline.Process($_)
        } catch {
            throw
        }
    }
    end
    {
        try {
            $steppablePipeline.End()
        } catch {
            throw
        }
    }
}

function logview ($app, $num) {
    # Add other log types
    # logview on its own should show all registered logs that can be
    if ($app -eq "choco") { Get-Content C:\ProgramData\chocolatey\logs\chocolatey.log -Last $num }
    # https://docs.microsoft.com/en-us/windows/deployment/upgrade/log-files
    # https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/windows-setup-log-files-and-event-logs
    # https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-logs
}

Set-Alias logs logview

function wh ($search, $path, $index, [switch]$size, [switch]$bare) { 
    # Could add a switch to tries --version, -version, --ver, -ver, --v, -v, /?  to see if get output from the program
    # Could add a switch to get internal version of executable:
    #    (Get-Item C:\Windows\System32\Lsasrv.dll).VersionInfo.FileVersion
    #    (Get-Command C:\Windows\System32\Lsasrv.dll).Version     # get updated (patched) ProductVersion by using this:
    # "original" vs "patched" is due to the way the FileVersion is calculated (see the docs here) https://stackoverflow.com/questions/30686/get-file-version-in-powershell

    if ([string]::IsNullOrWhiteSpace($search)) {
        ""
        "'wh' is a mix of 'which' from linux with 'where.exe' from Windows (and avoiding the 'where' PowerShell keyword)"
        ""
        "USAGE: wh <search> [<path>] [-index] [-size] [-bare]"
        "    <search> : String to search for (accepts wildcards)."
        "    <path>   : Optional path to search recursively instead of `$Env:PATH (default will search only on the PATH statement)."
        "    -index   : Optionally jump to (i.e. CD to) the path of a found item by the index in a 'wh' search."
        "    -size    : Optionally also show the file sizes (formatted as B/KB/MB/GB/TB)."
        "    -bare    : Optionally just return only the full paths (for further processing)."
        ""
        "    wh notepad.exe"
        "    wh notepad.exe 2"
        "    wh *.bat -bare   # Find every *.bat file on `$Env:PATH, just show the results only"
        "    wh *.bat 'C:\Program Files'   # Find every *.bat file under 'C:\Program Files'"
    }
    else {
        function Format-FileSize([int64]$size) {
            if ($size -gt 1TB) {[string]::Format("{0:0.00}TB", $size / 1TB)}
            elseif ($size -gt 1GB) {[string]::Format("{0:0.0}GB", $size / 1GB)}
            elseif ($size -gt 1MB) {[string]::Format("{0:0.0}MB", $size / 1MB)}
            elseif ($size -gt 1KB) {[string]::Format("{0:0.0}kB", $size / 1KB)}
            elseif ($size -gt 0) {[string]::Format("{0:0.0}B", $size)}
            else {""}
        }
        
        $count = 0
                                                                            # i(?:n(?:d(?:e(?:x)?)?)?)? or i(?:n(?:d(?:ex?)?)?)?) are equivalent
        if ([string]::IsNullOrWhiteSpace($path) -or $path -match "^\d+$" -or $path -match "\-(?:i(?:n(?:d(?:e(?:x)?)?)?)?)?") {    # In this case, we use $path as if it was $index (!), just pretend

            # If $path is not defined, or is an integer, then we are just looking at the PATH environment variable. i.e. using "where.exe"
            $results = where.exe $search   # case for Windows, could make this cross-platform with 'which' on Linux
            $total = $results.Count   # Count is number of elements always, but Length can vary as Legnth of a string etc.

            if ($path -notmatch "^\d+$") {
                if ($bare -ne $true) { "`nA total of $total matches for '$search' were found on `$Env:PATH`n" }

                foreach ($i in $results) {
                    $count++
                    $s = ""
                    if ($size) { $s = "$(Format-FileSize($i.Length)) :" ; $totalsize += $i.Length }
                    if ($bare -ne $true) { "Path $count : $s $i" } else { $i }
                    if ($path -eq $count) { $jumpto = split-path $i }
                }
                if ($bare -ne $true) { if ($size) { "`nTotal Size = $(Format-FileSize($totalsize))" } }
                if ($bare -ne $true) { "`nAppend one of the above index values to jump to (i.e. cd to) a found directory. e.g. 'wh *.exe 2' will cd to the 2nd found path.`nSpecify '-bare' to only show found result paths." }
            }
            else {
                foreach ($i in $results) {
                    $count++
                    if ($path -eq $count) { $jumpto = split-path $i }
                }
                Set-Location $jumpto
            }

            $count = 0
            if ($index -match "^\d+$") {   # now have to deal with the non-faking $path scenario if e.g. "-i 5" uses $index value properly
                foreach ($i in $results) {
                    $count++
                    if ($index -eq $count) { $jumpto = split-path $i }
                }
                Set-Location $jumpto
            }
        }
        else {   # case when $path can be a path is really a path, i.e. it is not nullorwhitespace, and it is not an integer

            if (Test-Path $path) {   # check if it really is a path, redefine $results as a recursive Get-ChildItem
                $results = Get-ChildItem -Path $path -Filter $search -Recurse -ErrorAction SilentlyContinue -Force
                $total = $results.Count   # Count is number of elements always, but Length can vary as Legnth of a string etc.
                $totalsize = 0

                if (([string]::IsNullOrWhiteSpace($index))) {
                    if ($bare -ne $true) { "`nA total of $total matches for '$search' were found inside '$path'`n" }

                    $count = 0
                    foreach ($i in $results) {
                        $count++
                        $s = ""
                        if ($size) { $s = "$(Format-FileSize($i.Length)) :" ; $totalsize += $i.Length }
                        if ($bare -ne $true) { "Path $count : $s $($i.FullName)" } else { $i.FullName }
                    }
                    if ($bare -ne $true) { if ($size) { "`nTotal Size = $(Format-FileSize($totalsize))" } }
                    if ($bare -ne $true) { "`nAppend one of the above index values to jump to (i.e. cd to) a found directory. e.g. 'wh *.exe 2' will cd to the 2nd found path.`nSpecify '-bare' to only show found result paths." }
                }
                else {
                    foreach ($i in $results.FullName) {
                        $count++
                        if ($index -eq $count) { $jumpto = split-path $i }
                    }
                    Set-Location $jumpto
                }

                $count = 0
                if ($index -match "^\d+$") {   # now have to deal with the non-faking $path scenario if e.g. "-i 5" uses $index value properly
                    foreach ($i in $results) {
                        $count++
                        if ($index -eq $count) { $jumpto = split-path $i }
                    }
                    Set-Location $jumpto
                }
            }
        }
    }
}

Set-Alias whereis wh

function whdups {
    # Extend "wh" to find all duplicates on system on all Path folders. Do this by checking for names that exist more than once.
    # Good example of Group-Object to take a count on the items
    # Tried Compare-Object but was a failure, Group-Object was better:   Compare-Object -ReferenceObject $files -DifferenceObject $files_sorted

    $paths = [Environment]::GetEnvironmentVariable("Path") -split ";" | select -unique | sort   # All paths (both Machine + User), remove duplcate paths and sort into an array
    $files = @()
    foreach ($path in $paths) {
        if (Test-Path $path) { $files += (gci $path -af | select Name).Name }   # Need ( ).Name so that we just have an array of strings. -af shorthand for -Attributes Files
    }
    
    # Find all NON-unique items in an array (tried various ways, but as usual, PowerShell has a very efficient way to do this)
    $dups = $files | Group-Object | Where-Object {$_.Count -gt 1}
    foreach ($d in ($dups | sort).Name) {
        "$d"
        wh $d -bare
        ""
    }

    # Another method using a hash instead of Group-Object that works well:
    #   $hash = @{}
    #   $files | foreach { $hash["$_"] += 1 }   # $hash["filename"] will increment for each found item
    #   $has.keys | where {$hash["$_"] -gt 1} | foreach { write-host "Duplicate element found $_" }   # Find keys that are more than 1 in size
    
    # Can use the above for any generic dedupe process, such as for media/films:
    # i.e. Collect a hash of all Names and put all paths into the values. Then compare the paths for sizes etc and present the output
}

Set-Alias which wh   # Might as well just alias 'which' to 'wh' in case type it while in PowerShell
Set-Alias which1 wh   # Might as well just alias 'which' to 'wh' in case type it while in PowerShell

function zip ($FilesAndOrFoldersToZip, $PathToDestination, [switch]$sevenzip, [switch]$maxcompress, [switch]$mincompress, [switch]$nocompress ) {
    # For most scenarios, it's most concise to simply test a variable (or expression) in an if-statement condition with no comparison
    # operators, as it covers variables that do not exist and $null values, as well as empty strings. For example:
    # if ($x) { <Variable "x" exists and is neither null nor contains an empty value> }
    # If you just want to know if a variable was never assigned a non-empty value, it is this simple:   if (-not $x)
    if (!$FilesAndOrFoldersToZip) {
        "Syntax: zip `FileMaskToZip [NameOfArchive] [-sevenzip] [-maxcompress] [-mincompress] [-nocompress]"
        "   By default, will try to find and use 7z.exe to create a .zip."
        "   Output archive gets date-time string by default. e.g. 'MyZip' => 'MyZip__2021-03-17__19-07-31.zip'."
        "   Compress-Archive will be used if 7z.exe is not found."
        "   -sevenzip will create a .7z archive instead of .zip"
        "   -maxcompress=-mx9 / -mincompress=-mx1 / -nocompress=-mx0"
        break
    }
    # zip a folder (by default recursively). Possibly add $password and [switch]$mincompress (-mx1) / $maxcompress (-mx9)
    # Always append date-time "2021-04-13__96_46_13" to every archive as means always unique and good for backups etc

    # Cannot create a switch with name "-7zip" as a switch : https://stackoverflow.com/questions/19006740/why-cant-parameter-names-start-with-a-number
    # but *can* have variables that start with a number oddly enough, so $7z is ok for a variable, just not for a switch(!)

    # <7z.exe Switches>           # -- : Stop switches and @listfile parsing
    # -ai[r[-|0]]{@listfile|!wildcard} : Include archives   # -ax[r[-|0]]{@listfile|!wildcard} : eXclude archives
    # -ao{a|s|t|u} : set Overwrite mode                     # -an : disable archive_name field
    # -bb[0-3] : set output log level                       # -bd : disable progress indicator
    # -bs{o|e|p}{0|1|2} : set output stream for output/error/progress line
    # -bt : show execution time statistics
    # -i[r[-|0]]{@listfile|!wildcard} : Include filenames
    # -m{Parameters} : set compression Method
    #   -mmt[N] : set number of CPU threads
    #   -mx[N] : set compression level: -mx1 (fastest) ... -mx9 (ultra)
    # -o{Directory} : set Output directory    # -p{Password} : set Password
    # -r[-|0] : Recurse subdirectories        # -sa{a|e|s} : set Archive name mode
    # -seml[.] : send archive by email
    # -slp : set Large Pages mode             # -snh : store hard links as links
    # -snl : store symbolic links as links    # -sni : store NT security information
    # -sns[-] : store NTFS alternate streams  # -so : write data to stdout
    # -spf : use fully qualified file paths
    # -stm{HexMask} : set CPU thread affinity mask (hexadecimal number)
    # -stx{Type} : exclude archive type       # -t{Type} : Set type of archive
    # -v{Size}[b|k|m|g] : Create volumes
    # -w[{path}] : assign Work directory. Empty path means a temporary directory
    # -x[r[-|0]]{@listfile|!wildcard} : eXclude filenames
    # -y : assume Yes on all queries

    # To recurse with the crappy PowerShell command, have to use -LiteralPath and path has to end with "\"
    # To recurse with 7z.exe, use -r[-|0] : Recurse subdirectories
    # -r Enable recurse subdirectories.
    # -r- Disable recurse subdirectories. This option is default for all commands.
    # -r0 Enable recurse subdirectories only for wildcard names.

    # 7z.exe a -y -aoa -p1234 -tzip xxx.zip ./*   # add to archive
    # 7z.exe x <filetounzip> -o./mypath -y -r   # extract
    # Need to split up each argument one by one (!!)

    $dtnow = Get-Date -format "yyyy-MM-dd__HH-mm-ss"

    # Just use wherever a copy of 7z.exe is found:
    $7z = ""
    if (Test-Path 'C:\Program Files (x86)\7-Zip\7z.exe') { $7z = 'C:\Program Files (x86)\7-Zip\7z.exe' }
    if (Test-Path 'C:\Program Files\7-Zip\7z.exe') { $7z = 'C:\Program Files\7-Zip\7z.exe' }
    if (Test-Path 'C:\Windows\7z.exe') { $7z = 'C:\Windows\7z.exe' }
    if (Test-Path 'D:\7-Zip\7z.exe') { $7z = 'D:\7-Zip\7z.exe' }
    if (Test-Path 'D:\0\7-Zip\7z.exe') { $7z = 'D:\0\7-Zip\7z.exe' }
    if ($7z -ne "") { "The 7z.exe found at '$7z' will be used.`n" }

    if ($maxcompress) { $compress = "-mx9"}
    if ($mincompress) { $compress = "-mx1"}
    if ($nocompress) { $compress = "-mx0"}

    # Strip any declared .7z or .zip to replace with correct timestamp and extension later
    if ($PathToDestination -match ".7z$") { $PathToDestination = $PathToDestination -replace ".7z$", "" }
    if ($PathToDestination -match ".zip$") { $PathToDestination = $PathToDestination -replace ".zip$", "" }

    if ($sevenzip) {
        $PathToDestination = "$($PathToDestination)__$($dtnow).7z"
        if ($7z -ne "") {
            echo "$7z a -y -r -t7z $compress $PathToDestination $FilesAndOrFoldersToZip"
            & $7z "a" "-y" "-r" "-t7z" "$compress" "$PathToDestination" "$FilesAndOrFoldersToZip"
        }
    }
    else {
        $PathToDestination = "$($PathToDestination)__$($dtnow).zip"
        if ($7z -ne "") {
            echo "$7z a -y -r -t7zip $compress $PathToDestination $FilesAndOrFoldersToZip"
            & $7z "a" "-y" "-r" "-tzip" "$compress" "$PathToDestination" "$FilesAndOrFoldersToZip" 
        }
        else { 
            # Don't really want to use PowerShell method but will if 7z.exe is not available
            if ($FilesAndOrFoldersToZip -match '\\$') {
                # Use -LiteralPath with "\" at end of line to recurve whole folder with Compress-Archive
                Compress-Archive -LiteralPath $FilesAndOrFoldersToZip -DestinationPath $PathToDestination
            }
            else {
                Compress-Archive -Path $FilesAndOrFoldersToZip -DestinationPath $PathToDestination
            }
        }
    }
}   

# https://blogs.msmvps.com/russel/2017/04/28/resizing-the-powershell-console/
function lll {
    # Adjust console window position, half-screen, left side
    Add-Type -AssemblyName System.Windows.Forms
    $Screen = [System.Windows.Forms.Screen]::PrimaryScreen
    $width = $Screen.WorkingArea.Width   # .WorkingArea ignores the taskbar, .Bounds is whole screen
    $height = $Screen.WorkingArea.Height
    $w = $width/2 + 13
    $h = $height + 8
    $x = -7
    $y = 0
    Set-WindowNormal
    Set-ConsolePosition $x $y $w $h

    $MyBuffer = $Host.UI.RawUI.BufferSize
    $MyWindow = $Host.UI.RawUI.WindowSize
    $MyBuffer.Height = 9999
    "`nWindowSize $($MyWindow.Width)x$($MyWindow.Height) (Buffer $($MyBuffer.Width)x$($MyBuffer.Height))"
    "Position : Left:$x Top:$y Width:$w Height:$h (pixels)`n"
}

function rrr {
    # Adjust console window position, half-screen, right side
    Add-Type -AssemblyName System.Windows.Forms
    $Screen = [System.Windows.Forms.Screen]::PrimaryScreen
    $width = $Screen.WorkingArea.Width   # .WorkingArea ignores the taskbar, .Bounds is whole screen
    $height = $Screen.WorkingArea.Height
    $w = $width/2 + 13
    $h = $height + 8
    $x = $w - 20
    $y = 0
    Set-WindowNormal
    Set-ConsolePosition $x $y $w $h

    $MyBuffer = $Host.UI.RawUI.BufferSize
    $MyWindow = $Host.UI.RawUI.WindowSize
    $MyBuffer.Height = 9999
    "`nWindowSize $($MyWindow.Width)x$($MyWindow.Height) (Buffer $($MyBuffer.Width)x$($MyBuffer.Height))"
    "Position : Left:$x Top:$y Width:$w Height:$h (pixels)`n"
}
function fff {
    # Adjust console window position, full size (non-maximised)
    Set-ConsolePosition -7 0 600 600
    if ($Host.Name -match "console") {
        $MaxWidth  = $host.UI.RawUI.MaxPhysicalWindowSize.Width
        $MaxHeight = $host.UI.RawUI.MaxPhysicalWindowSize.Height - 1
        $MyBuffer  = $Host.UI.RawUI.BufferSize
        $MyWindow  = $Host.UI.RawUI.WindowSize

        Set-WindowNormal
        $MyWindow.Height = $MaxHeight
        $MyWindow.Width = $Maxwidth
        $MyBuffer.Height = 9999
        $MyBuffer.Width = $Maxwidth
        # $host.UI.RawUI.set_bufferSize($MyBuffer)
        # $host.UI.RawUI.set_windowSize($MyWindow)
        $host.UI.RawUI.BufferSize = $MyBuffer
        $host.UI.RawUI.WindowSize = $MyWindow
        "`nWindowSize $($MyWindow.Width)x$($MyWindow.Height) (Buffer $($MyBuffer.Width)x$($MyBuffer.Height))`nPosition : Left:-7 Top:0 Width:$MaxWidth Height:$MaxHeight (pixels)"
    }
}

function mmm {
    # mmm for "maximize"
    Set-WindowMax
}

function ccc {
    # Adjust console window position, centre top
    Set-ConsolePosition -7 25 600 600
    if ($Host.Name -match "console") {
        Set-ConsolePosition 75 25 600 600
        Set-WindowNormal
        Set-MaxWindowSize
    }
}

# Template for future use ... can be handy 
function Move-Mouse {
    # Load Required Assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Cursor]::Position   # Get current location
    $Position = [System.Drawing.Point]::new(120,220)   # Move to 120,220
    [System.Windows.Forms.Cursor]::Position = $Position
}

# Remove comments and indentation from code.
# i.e. make a flat file that is still functional, but without comments
# Template, this is just for WOTR Toolkit Autohotkey, but useful techniques for elsewhere
function Remove-Comments {
    $x = Get-Content ".\WOTR Toolkit.ahk"   # $x is an array, not a string
    $out = ".\WOTR Toolkit 1.ahk"
    if (Test-Path $out) { Clear-Content $out }   # wipe the output file

    $x = $x -split "[\r\n]+"               # Remove all consecutive line-breaks, in any format '-split "\r?\n|\r"' would just do line by line
    $x = $x | ? { $_ -notmatch "^\s*$" }   # Remove empty lines
    $x = $x | ? { $_ -notmatch "^\s*;" }   # Remove all lines starting with ; including with whitespace before
    $x = $x | % { ($_ -split " ;")[0] }    # Remove end of line comments
    $x = ($x -replace $regex).Trim().Trim()
    Set-Content $out $x
    # $x | Out-File $out
    $x | more
}

####################
#
# Setup various optional components
#
# Don't want to completely automate these installations
# Main install should just focus on essential components:
# - Latest PowerShell + Chocolatey + Boxstarter
# - Various useful modules, various useful scripts
# - Collecting core tools into the Profile Extensions
#
# Everything else should be in the Custom-Tools Module.
#
####################

function IfExistSkipCommand ($toCheck, $toRun) {
    if (Test-Path($toCheck)) {
        Write-Host "Item exists        : $toCheck" -ForegroundColor Green
        Write-Host "Will skip installer: $toRun`n" -ForegroundColor Cyan
    } else {
        Write-Host "Item does not exist: $toCheck" -ForegroundColor Green
        Write-Host "Will run installer : $toRun`n" -ForegroundColor Cyan
        Invoke-Expression $toRun

    }
}

function Install-WindowsTerminal {
    # download installation code
    $code = Invoke-WebRequest -Uri 'https://chocolatey.org/install.ps1' -UseBasicParsing
    # invoke installation code
    Invoke-Expression $code
    # install windows terminal
    choco install microsoft-windows-terminal -y
    # Setting up console colours ...
    # https://thecustomizewindows.com/2020/08/windows-terminal-with-solarized-dark-oh-my-zsh-theme/
    # https://www.reddit.com/r/Windows10/comments/4jbguv/changing_linux_terminal_colors_to_solarized_theme/
}

function Install-Random {
"
Placeholder for interesting / useful / fun things to install.
cinst dwarf-fortress -y   # dwarf-fortress 0.47.04, installs to C:\opt\df
Should probably look at redirecting this somwhere, say to `"D:\Dwarf Fortress`" or something.
"
}

function Install-NotepadPlusPlus {
    <#
    .SYNOPSIS
    Install Notepad++ with customised settings. If Notepad++ is installed, will uninstall, then reinstall latest chocolatey package.
    This is useful as can control updates more easily compared to the default installers (instead of annoying "do you want to upgrade?" prompts).
    .LINK
    https://notepad-plus-plus.org/   # Home page
    https://chocolatey.org/packages/notepadplusplus   Chocolatey package
    https://medium.com/@HiSandy/notepad-plus-plus-material-theme-2c3951e65e01   Customise Notepad++
    #>

    if ((New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator) -ne $true) {
        "Chocolatey operations require Administrator elevation.`nTry 'sudo <command>   , or   sudo { <command1> ; <command2> ; ... }`n"
        break
    }
    ""
    # First, test if notepad++.exe is currently running, and if so prompt to close it
    while (ps *notepad++*) {
        read-host "Cannot contiune while Notepad++ is open, please close manually then continue, or Ctrl-C to quit this function."
        # $updateonline = read-host "No local file found, download latest Profile Extensions from internet (default is y) (y/n)? "
    }

    "Notepad++ is closed so setup will now continue (do not reopen notepad++ until setup completes)."

    # Test if choco package is installed, if not, then it must have been installed manually, perform uninstall
    # But *only* do this if it was not installed by chocolatey as it does install to the default location
    # $parts = $result.Split(' ')   #  Optionally return version : return $parts[1] -eq $this.Version.
    # $parts[1]   # contains the version number of the package
    $result = C:\ProgramData\Chocolatey\choco list -lo | Where-object { $_.ToLower().StartsWith("NotepadPlusPlus".ToLower()) }
    if ($result -ne $null) {
        # Check if C:\Program Files\Notepad++\notepad++.exe exists, if not, need to uninstall then reinstall
        if (!(Test-Path "C:\Program Files\Notepad++\notepad++.exe")) {
            choco uninstall -y notepadplusplus notepadplusplus.install --force
            choco install -y notepadplusplus notepadplusplus.install --force
        }
        else {
            choco upgrade -y notepadplusplus.install notepadplusplus   # Note that 'update' is a deprecated term, always use 'upgrade'
        }
        # see what latest package version is, if so update to it
        # The notepad++ package is installed, so attempt an update

        # choco messages when trying to uninstall:
        # cuninst notepadplusplus : You are uninstalling notepadplusplus, which is likely a metapackage for an *.install/*.portable package that it installed (notepadplusplus represents discoverability).
        # cuninst notepadplusplus.install : notepadplusplus.install not uninstalled. An error occurred during uninstall: Unable to uninstall 'notepadplusplus.install 7.8.5' because 'notepadplusplus 7.8.5' depends on it.
        # So, must uninstall notepadplusplus then the .install package in that order.
    }
    else
    {
        "No chocolatey Notepad++ package was found, will remove versions if present then reinstall`n"
        # If choco packages are not there, safe to attempt silent uninstall by notepad++ uninstaller
        # Easiest way to run Cmd commands are to use here-strings (to freely use " inside) and then Invoke-Expression (iex) against that
        # https://social.technet.microsoft.com/Forums/ie/en-US/7b398cea-0d29-4588-a6bc-ef793b51cc3c/run-a-dos-command-in-powershell?forum=winserverpowershell
        # As far as I'm aware the silent uninstall option was broken due to the prompt about keeping custom settings. When this was fixed in v7.2 I believe the fix was to leave those files by default, in the program files location. As you are deploying a new version of the same program, it would seem a bad decision if you were to also remove your end users customisations.
        # start /wait npp.7.2.Installer.x64.exe /S
        $command = @'
cmd.exe /C "C:\Program Files (x86)\Notepad++\uninstall.exe" /S
'@ ; if (Test-Path "C:\Program Files (x86)\Notepad++\uninstall.exe") { Invoke-Expression -Command:$command }
        
        $command = @'
cmd.exe /C "C:\Program Files\Notepad++\uninstall.exe" /S
'@ ; if (Test-Path "C:\Program Files\Notepad++\uninstall.exe") { Invoke-Expression -Command:$command }

        # Now install the latest chocolatey package. Note: installs to C:\Program Files. i.e. not to the "(x86)" folder.
        choco install -y notepadplusplus --force
    }

    "Apply registry fixes to configure Notepad++ to known state:"   # Following assumes 64-bit version for all configuration
    "- Must open and then close Notepad++ for config.xml to be generated (cannot make config changes without this)"
    "- Let Notepad++ open and then close to complete setup"
    # $name = "notepad++.exe"   # For some reason needed the full "notepad++.exe" even though "notepad++" works fine when not passing a variable
    Start-Process "notepad++.exe"       # This will use the chocolatey shim in C:\ProgramData\Chocolatey\bin
    # $status = Get-Process "notepad++" -EA silent   # Note that process name does *not* include .exe
    # while (!($status)) { "Waiting for notepad++ to start" ; Start-Sleep -Seconds 1 }
    # Gracefully shutdown process   https://stackoverflow.com/questions/28481811/how-to-correctly-check-if-a-process-is-running-and-stop-it
    # This is important as killing Notepad++ does not create config.xml required to update settings.
    # Might not actually need this, I might just be closing Notepad++ too quickly
    Sleep 2   # Leave open for at least 2 seconds
    $npp_process = Get-Process "notepad++" -ErrorAction SilentlyContinue
    if ($npp_process) {
          $npp_process.CloseMainWindow() | Out-Null   # try graceful application close first
          Sleep 2                                                               # Allow 2 seconds for process to stop
          if (!$npp_process.HasExited) { $npp_process | Stop-Process -Force }   # If it did not exist cleanly, stop app with -Force
    }
    # Remove-Variable npp_process
    # Stop-Process -Name "notepad++"

    # Registry change to make Notepad++ the default text editor. i.e. essentially replaces notepad by Notepad++
    # Make the notepad++ version in C:\Program Files the default handler for all notepad text file requests
    # x86 : cmd.exe /C reg add "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" /v "Debugger" /t REG_SZ /d "\"%ProgramFiles(x86)%\Notepad++\notepad++.exe\" -notepadStyleCmdline -z" /f
    # x64 : cmd.exe /C reg add "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" /v "Debugger" /t REG_SZ /d "\"%ProgramFiles%\Notepad++\notepad++.exe\" -notepadStyleCmdline -z" /f
    # cmd alternative syntax with reg add: reg add "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" /v "Debugger" /t REG_SZ /d "\"%ProgramFiles(x86)%\Notepad++\notepad++.exe\" -notepadStyleCmdline -z" /f
    
    # Not sure why I was testing for the existence of notepad.exe, I think should always force the insertion of this 
    # if (Test-Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe") {
    "- Override Default Text Editor (Notepad -> Notepad++) by updating Registry"
    New-Item -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" -Name "Debugger" -Value "`"C:\Program Files\Notepad++\notepad++.exe`" -notepadStyleCmdline -z" -Force
    # }

    # Make Notepad++ the default association for files with no extension (such as "Dockerfile")
    # Even though VSCode will be used for editing those, Notepad++ should be default for all text file types.    
    # assoc .txt="Notepad"
    # ftype "Notepad"="C:\Program Files (x86)\Notepad++\notepad++.exe" "%1"
    # https://www.file-extensions.org/article/set-default-program-for-files-with-no-extension-in-windows
    # https://www.foxinfotech.in/2018/06/how-to-associate-a-file-type-with-language-in-notepad.html
    # https://github.com/notepad-plus-plus/notepad-plus-plus/issues/1786
    "- Configure no-extension files to default to Notepad++"
    $command = @'
    cmd.exe /C assoc .=""No Extension""
'@ ; Invoke-Expression -Command:$command | Out-Null
    $command = @'
    cmd.exe /C ftype ""No Extension""=""C:\Program Files\Notepad++\notepad++.exe"" ""%1""
'@ ; Invoke-Expression -Command:$command | Out-Null
    
    # Main app customisations:
    # - turn off updates (chocolatey will handle updates via "cup all" to prevent the annoying "do you want to upgrade?" prompts)
    # - replace tabs by spaces
    # - turn off remember sessions
    # - disable auto-completion (i.e. just use Notepad++ as a text editor, will mainly use programming editors like VS Code or ISE for programming)
    # - enable Word Wrap by default (VS Code is main scripting editor, Notepad++ used for text files, so usually always want Word Wrap)
    $npp_config = "$env:AppData\Notepad++\config.xml"
    if (!(Test-Path $npp_config)) { "config.xml is missing, open Notepad++ and close again to create this file" ; pause }
    $xml = New-Object XML   # Always declare XML object first as opposed to other constructs, this is around 7x faster performance
    $xml.Load($npp_config)
    # Keep all the formats (tabs and spaces etc) the same. The Property name is confusing.
    $xml.PreserveWhitespace = $true
    "- Change NoUpdate value to 'yes' (disable auto-updates)"
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "noupdate"}).'#text' = 'yes'              # item is "noupdate" so 'yes' turns off updates!
    "- Replace all tabs by spaces"
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "TabSetting" }).replaceBySpace = 'yes'    # Replace all tabs by spaces
    "- Don't reopen last-opened tabs at every start"
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "RememberLastSession" }).'#text' = 'no'   # Don't reopen last open tabs every time
    "- Turn off auto-completion (more a scripting tool function, not useful for a basic text editor)"
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "auto-completion" }).'autoCAction' = '0'  # Turn off auto-completion
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "auto-completion" }).'triggerFromNbChar' = '1'    # ???
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "auto-completion" }).'autoCIgnoreNumbers' = 'no'  # ???
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "auto-completion" }).'funcParams' = 'no'          # function parameters hint on input
    "- Turn on line wrap (focus is as text editor rather than programming so line wrap is useful)"
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "ScintillaPrimaryView" }).'Wrap' = 'yes'          # Turn on line wrap
    "- Prevent the new default of not giving a .txt extension when saving files"
    ($xml.NotepadPlus.GUIConfigs.GUIConfig | where { $_.name -eq "MISC" }).'newStyleSaveDlg' = 'no'        # Leave .txt as the default save-fille extension
    # contents of the xml file at $env:AppData\Notepad++\config.xml
    #    <GUIConfig name="noUpdate" intervalDays="15" nextUpdateDate="20191126">no</GUIConfig>
    #    <GUIConfig name="TabSetting" replaceBySpace="yes" size="4" />
    #    <GUIConfig name="AppPosition" x="0" y="0" width="1100" height="700" isMaximized="yes" />
    #    <GUIConfig name="RememberLastSession">yes</GUIConfig>
    #    <GUIConfig name="auto-completion" autoCAction="3" triggerFromNbChar="1" autoCIgnoreNumbers="no" funcParams="yes" />
    #    <GUIConfig name="ScintillaPrimaryView" lineNumberMargin="show" bookMarkMargin="show" indentGuideLine="show" folderMarkStyle="box" lineWrapMethod="aligned" 
    #      currentLineHilitingShow="show" scrollBeyondLastLine="no" disableAdvancedScrolling="no" wrapSymbolShow="hide" Wrap="yes" borderEdge="yes" edge="no"
    #      edgeNbColumn="80" zoom="0" zoom2="0" whiteSpaceShow="hide" eolShow="hide" borderWidth="2" smoothFont="no" />
    $xml.Save($npp_config)

    ""
    "To keep Notepad++ up to date:   cup -y notepadplusplus      or,     cup -y all"
    ""
}

function Install-Steam ($SetupFolder) {
    "Not implemented yet ..."
    # https://support.steampowered.com/kb_article.php?ref=7418-YUBN-8129
    # The following instructions are a simple way to move your Steam installation along with your games:
    # Exit the Steam client application.
    # Browse to the Steam installation folder for the Steam installation you would like to move (C:\Program Files\Steam by default).
    # Delete all of the files and folders except the SteamApps & Userdata folders and Steam.exe
    # Cut and paste the whole Steam folder to the new location, for example: D:\Games\Steam\
    # Launch Steam and log into your account.
    # Steam will briefly update and then you will be logged into your account. For installed games, verify your game cache files and you will be ready to play. All future game content will be downloaded to the new folder (D:\Games\Steam\SteamApps\ in this example)
    # If the updating fails, rerunning the Steam installer pointing at the new location correct that.
    # 
    # Best approacb might be:
    # test for Steam install in usual C: locations.
    # If there, move the SteamApps folder to the chosen folder
    # Install Steam to that folder so that the Steam games are then 
}

function Install-ConsoleColors {
    # None of this works properly, needs work!
    Write-Host "concfg notes: https://stackoverflow.com/questions/13690223/how-can-i-launch-powershell-exe-with-the-default-colours-from-the-powershell-s"
        Write-Host "Various good methods here!"
        Write-Host "cd hkcu:/console"
        Write-Host "$0 = '%systemroot%_system32_windowspowershell_v1.0_powershell.exe'"
        Write-Host "ni $0 -f"
        Write-Host "sp $0 ColorTable00 0x00562401"
        Write-Host "sp $0 ColorTable07 0x00f0edee"
}

# https://www.comparitech.com/net-admin/powershell-cheat-sheet/

function Help-RandomGuides {
    "
https://mangolassi.it/topic/22381/setup-nextcloud-19-0-4-on-fedora-32   https://mangolassi.it/category/2/it-discussion
https://mangolassi.it/topic/12603/opinions-ansible-vs-saltstack/4
https://chocolatey.org/resources/case-studies/automated-managing-windows-desktop-software
https://chocolatey.org/docs/features-create-packages-from-installers

"
}

function Help-Git {
    "
Quick notes on basic git operations:

To CONFIGURE:
git config --global user.email 'roysubs@hotmail.com'
git config --global user.name 'Roy Subs'

To DOWNLOAD:
git clone https://github.com/roysubs/custom-tools.git   # Will clone into a subfolder of the currect folder with name 'custom-tools'
git clone https://roysubs:<GITHUB_ACCESS_TOKEN>@github.com/roysubs/custom-tools.git   # Clone with full access rights (so git push works automatically)

To UPLOAD (always add / commit first!)
git add .     # Add all files
git commit    # On Windows this requires a message
git reset --soft HEAD~1   # Undo to HEAD minus 1
git push https://<GITHUB_ACCESS_TOKEN>@github.com/roysubs/custom-tools.git

Enumerating objects: 11, done.
Counting objects: 100% (11/11), done.
Delta compression using up to 4 threads
Compressing objects: 100% (6/6), done.
Writing objects: 100% (6/6), 2.80 KiB | 409.00 KiB/s, done.
Total 6 (delta 4), reused 0 (delta 0), pack-reused 0
remote: Resolving deltas: 100% (4/4), completed with 4 local objects.
remote: This repository moved. Please use the new location:
remote:   https://github.com/roysubs/Custom-Tools.git
To https://github.com/roysubs/custom-tools.git
   d1df5b1..0a88586  main -> main
"
}


function Help-robocopy {
    "
To duplicate a folder:
robocopy . D:\Backup\`$(Get-Date -Format yyyyMMdd_HHmm)_Robocopy_Backup /mir /r:1 /w:1 /create
/mir 'mirror', /r:1 'retry 1 time', /w:1 'wait 1 sec between retry', /create 'create zero-size files only'

robocopy . D:\Backup\`$(Get-Date -Format yyyyMMdd_HHmm)_Robocopy_Backup /mir /create /copy:DAT /DCOPY:T
/COPY:copyflag[s] :: What to copy for files (default is /COPY:DAT) : D=Data, A=Attributes, T=Timestamps, S=Security=NTFS ACLs, O=Owner info, U=aUditing info).
/DCOPY:T :: Copy Directory Timestamps.

Without using /mir
robocopy . D:\Backup\`$(Get-Date -Format yyyyMMdd_HHmm)_Robocopy_Backup /COPYALL /E /SECFIX /DCOPY:T
/E copies empty folders (remove if not needed)
/SECFIX copy NTFS permissions (remove if not needed)
/XO can be added to exclude older (ie if doing a true-up for a folder migration)
"
}

function Help-Chrome {
    "
Duplicate Tab Helper: If a tab already exists for a given URL, prevent it opening a 2nd time
https://chrome.google.com/webstore/detail/duplicate-tab-helper/oaceoebbkmkgfjhmngdinoclnionlgoh?hl=en-GB

Save entire page as image (including scrolling whole page)
???

Cookie Keeper: Set domain rules for cookies you want to keep, all other cookies are deleted.
https://chrome.google.com/webstore/detail/cookie-keeper/hkdjopjjoogbicnmbcenniplfnmcnhof
"
}

function Help-ToolkitCoreApps {
    ""
    "####################"
    "#"
    "# Selected Essential Apps useful for most systems:"
    "#"
    "####################"
    ""
    "Some have customised settings using Chocolatey installation with updated settings and make silent updates"
    "easy to manage. Use 'def' to check details: e.g.   'def Install-Notepad++'"
    ""
    "   Install-WindowsTerminal   # Also: Help-WindowsTerminal"
    "   Install-7zip"
    "   Install-Notepad++"
    "   Install-SumatraPDF"
    "   Install-Ditto"
    "   Install-VSCodePortable"
    "   Install-BitCometPortable"
    "   Install-WinSCP"
    "   Install-Firefox"
    "   Install-Chrome"
    "   Install-EdgeChromium"
    "   Install-FileZilla"
    "   Install-FileZilla-Server"
    "   Install-Macrium-Reflect"
    "   Install-AnyDesk"
    "   Install-Python"
    "   Install-Java"
    "   Install-AutoHotkey"
    "   Install-Steam"
    ""
    "Selected lists of additional Chocolatey packages:"
    ""
    (cat $HomeFix\Documents\WindowsPowerShell\Modules\Custom-Tools\Custom-Tools.psm1 | sls "function Help-Choco" | sls '"function Help-Choco"' -NotMatch).Line -replace " {", "" -replace "function ", "   "
    ""
}

function Help-WindowsTerminal {
    "wt.exe [options] [commands] to open a new instance of Windows Terminal"
    "[options]  : --help [-h, -?, /?] , --maximized [-M], --fullscreen [-F]"
    "[commands] : new-tab [nt], --profile [-p] <profile-name>, --startingDirectory [-d] <starting-directory>, --title"
    "   split-pane [sp], -h [--horizontal], -v [--vertical], --profile [-p] <profile-name>, -d starting-directory, commandline, --title"
    "   focus-tab [ft], --target [-t] <tab-index> Focuses on a specific tab."
    ""
    '   wt -p "Ubuntu-18.04"   # Open Ubuntu WSL'
    "   wt -d d:\              # Open new tab in D:\"
    '   wt `; `; `;  (from PowerShell),   wt ; ; ;  (from DOS),   cmd.exe /c "wt.exe" \; \;  (from Linux)   # Open three tabs'
    '   wt -p "Command Prompt" `; new-tab -p "Windows PowerShell"'
    '   wt -p "Command Prompt" `; split-pane -p "Windows PowerShell" `; split-pane -H wsl.exe'
    '   wt -p "Command Prompt" `; split-pane -V wsl.exe `; new-tab -d c:\ `; split-pane -H -d c:\ wsl.exe'
    '   wt --title tabname1 `; new-tab -p "Ubuntu-18.04" --title tabname2'
    '   wt `; new-tab -p "Ubuntu-18.04" `; focus-tab -t 1'
    "   start wt 'new-tab `"cmd`" ; split-pane -p `"Windows PowerShell`" ; split-pane -H wsl.exe'"
    ""
    "SSH Consoles: https://stackoverflow.com/questions/57363597/how-to-use-a-new-windows-terminal-app-for-ssh"
    "   https://docs.microsoft.com/en-us/windows/terminal/tutorials/ssh"
    "Powerline for Git Intergration: https://docs.microsoft.com/en-us/windows/terminal/tutorials/powerline-setup"
    "Tips and Tricks: https://docs.microsoft.com/en-us/windows/terminal/tips-and-tricks"
    "Cascadia Code Font: https://docs.microsoft.com/en-us/windows/terminal/cascadia-code"
    ""
    ""
    "Creating a new pane to the right of the focused pane (vertical split). Alt+Shift+= (equals)"
    "Creating a new pane below the focused pane (horizontal split).         Alt+Shift+- (minus)"
    "Key bindings, use the 'splitPane' action and 'vertical' 'horizontal' 'auto' values for the split property in your profiles.json file."
    "Note that the 'auto' value splits in the direction that has the longest edge to create a pane."
    "Console tabs can be duplicated with all information by a 'splitMode' property with 'duplicate' as the value to the splitPane key binding."
}

function Install-WindowsTerminal {
    "Add latest Windows Terminal package using Chocolatey"
    cup microsoft-windows-terminal -y
}

function Install-VSCodePortable {
    <#
    .SYNOPSIS
    Ideally, want a way to create an up to date VS Code Portable package.
    i.e. Get the Portable edition, update it with all required extensions, then zip that copy ready to take onto an offline server.
    #>
    # Official Portable method: https://code.visualstudio.com/docs/editor/portable
    # This mode enables all data created and maintained by VS Code to live near itself, so it can be moved around across environments.
    # This also provides a way to set the installation folder location for VS Code extensions (to avoid restrictions on access to the AppData folder).
    # Just unzip the VS Code download, then create a data folder within VS Code's folder:
    # By doing this, from then on, that folder will be used to contain *all* VS Code data, including session state, preferences, extensions, etc.
    # This data folder can also be moved to another VS Code installations so is very useful to updating a portable VS Code version.
    # Direct .zip (64-bit) link : https://go.microsoft.com/fwlink/?Linkid=850641
    # From this page https://code.visualstudio.com/docs/?dv=winzip

    # Could either, download the latest official portable version, or install the Chocolatey (portable) package. # cinst vscode.portable
    # https://www.raymond.cc/blog/"
    # Automate Trial-Resets: https://www.raymond.cc/blog/how-to-extend-the-trial-period-of-a-software/"

    # Chocolatey VSCode related packages:
    # cinst vscode -y   # 1.40.0 [Approved]    # vscode.portable 1.40.0 [Approved] Downloads cached for licensed users
    # cinst visualstudiocode-insiders --pre    # vscode insider, gets beta updates
    # cinst vscodium -y   # 1.40.0 https://ar.al/2019/10/24/how-to-migrate-from-vscode-to-vscodium-the-best-code-editor-ever-minus-the-corporate-bullshit/
    #    VS Codium is a fork of VS Code without Microsoft telemetry. In Linux this has updates maintained on apt, but not on Windows so updates don't work inside, use chocolatey to update
    # VisualStudioCode 1.23.1.20180730 [Approved]
    # chocolatey-vscode 0.7.1 [Approved] # chocolatey-vscode.extension 1.1.0 [Approved]
    #    This extension brings support for Chocolatey to Visual Studio Code.
    #    Chocolatey: Create new Chocolatey package to create the default templated Chocolatey package at the root of current workspace.
    #    Chocolatey: Pack Chocolatey package(s) to search current workspace for nuspec files and package them
    #    Chocolatey: Delete Chocolatey package(s) to search current workspace for nupkg files and delete them
    #    Chocolatey: Push Chocolatey package(s) to search current workspace for nupkg files and push them
    # cinst visualstudiocode-disableautoupdate -y   # Turn off auto-updates (useful if want chocolatey to control all updates)
    # cinst openinvscode 1.3.19 [Approved] Downloads cached for licensed users
}

function Install-VSCodeExtras {
    "
ToDo: Using Git, Using side-by-side diff, etc ...

Open empty temp files with preset extensions in VS Code to prompt it to install syntax extensions.
Useful so these are available when working offline

    Temp Ansible YAML.yml
    Temp ASP.NET.asp
    Temp Autohotkey.ahk
    Temp Batch.bat
    Temp Batch.cmd
    Temp C-Sharp.cs
    Temp Dockerfile
    Temp HTML.htm
    Temp HTML.html
    Temp Ini config files.ini
    Temp JavaScript.js
    Temp LaTeX.tex
    Temp PowerShell.ps1
    Temp Python.py
    Temp Registry.reg
    Temp SQL.sql
    Temp Windows Scripting Host.wsh
    Temp XML.xml
" | more

    function Touch ($file)
    {
        if ($file -eq $null) { throw "No filename supplied" }
        if (Test-Path $file) { (Get-ChildItem $file).LastWriteTime = Get-Date }
        else { New-Item -ItemType File $file }
    }

    # Create temp files with the extension syntax required
    touch "$env:TEMP\Temp Ansible YAML.yml"
    touch "$env:TEMP\Temp ASP.NET.asp"
    touch "$env:TEMP\Temp Autohotkey.ahk"
    touch "$env:TEMP\Temp Batch.bat"
    touch "$env:TEMP\Temp Batch.cmd"
    touch "$env:TEMP\Temp C-Sharp.cs"
    touch "$env:TEMP\Temp Dockerfile"
    touch "$env:TEMP\Temp HTML.htm"
    touch "$env:TEMP\Temp HTML.html"
    touch "$env:TEMP\Temp Ini config files.ini"
    touch "$env:TEMP\Temp JavaScript.js"
    touch "$env:TEMP\Temp LaTeX.tex"
    touch "$env:TEMP\Temp PowerShell.ps1"
    touch "$env:TEMP\Temp Python.py"
    touch "$env:TEMP\Temp Registry.reg"
    touch "$env:TEMP\Temp SQL.sql"
    touch "$env:TEMP\Temp Windows Scripting Host.wsh"
    touch "$env:TEMP\Temp XML.xml"

    # Open the temp files in VS Code (while online) to prompt extension installation
    code "$env:TEMP\Temp Ansible YAML.yml"
    code "$env:TEMP\Temp ASP.NET.asp"
    code "$env:TEMP\Temp Autohotkey.ahk"
    code "$env:TEMP\Temp Batch.bat"
    code "$env:TEMP\Temp Batch.cmd"
    code "$env:TEMP\Temp C-Sharp.cs"
    code "$env:TEMP\Temp Dockerfile"
    code "$env:TEMP\Temp HTML.htm"
    code "$env:TEMP\Temp HTML.html"
    code "$env:TEMP\Temp Ini config files.ini"
    code "$env:TEMP\Temp JavaScript.js"
    code "$env:TEMP\Temp LaTeX.tex"
    code "$env:TEMP\Temp PowerShell.ps1"
    code "$env:TEMP\Temp Python.py"
    code "$env:TEMP\Temp Registry.reg"
    code "$env:TEMP\Temp SQL.sql"
    code "$env:TEMP\Temp Windows Scripting Host.wsh"
    code "$env:TEMP\Temp XML.xml"

    ### Could use chocolatey packages for these, which would mean could auto-update all with 'cup -y'
    # cinst vscode-powershell -y       # 2019.11.0 [Approved]
    # cinst vscode-ansible -y          # 0.5.2 [Approved]
    # cinst vscode-csharp -y           # 1.21.7 [Approved]
    # cinst vscode-csharpextensions -y # 1.0.0.20180620 [Approved]
    # cinst vscode-python -y           # 2019.10.44104 [Approved]
    # cinst vscode-docker -y           # 1.0.0.20190907 [Approved]
    # cinst vscode-puppet -y           # 0.21.0 [Approved]
    # cinst vscode-ruby -y             # 0.25.3 [Approved]
    # cinst vscode-autohotkey -y       # 0.2.2 [Approved]
    # cinst vscode-java -y             # 0.52.0 [Approved]
    # cinst vscode-yaml -y             # 0.5.3 [Approved]
    # cinst vscode-codespellchecker -y # 1.0.0.20181011 [Approved]
    # cinst vscode-gitlens -y          # 1.0.0.20181011 [Approved]
    # cinst vscode-mssql -y            # 1.6.0 [Approved]

    # cinst vscode-settingssync -y     # 1.0.0.0 [Approved]

    # There are other extensions in chocolately, but the above seem generally useful to pre-install
    # there is also a way to install extensions using 'code.exe' 
    # You can see the full name for extensions from: C:\Users\<user>\.vscode\extensions
    # or by going to extensions manager in VS Code and searching for an extension
    # code --list-extensions
    # code --install-extension ms-vscode.csharp
    # code --uninstall-extension ms-vscode.csharp
    
    # code --install-extension ms-vscode.csharp
    # code --install-extension ms-vscode.powershell
    # code --install-extension ms-python.python
    # code --install-extension ms-azuretools.vscode-docker
    # code --install-extension ms-mssql.mssql
    # code --install-extension James-Yu.latex-workshop
    # code --install-extension slevesque.vscode-autohotkey
    # code --install-extension --force ms-vscode.powershell
    # code --install-extension --force ms-vscode-remote.remote-wsl
    # code --install-extension --force slevesque.vscode-autohotkey

    # https://stackoverflow.com/questions/34286515/how-to-install-visual-studio-code-extensions-from-command-line
    # code --install-extension --force DavidWang.ini-for-vscode
    # code --install-extension --force ionutvmi.reg
    # code --install-extension --force ms-vscode.cpptools --force
    # code --install-extension --force ritwickdey.liveserver
    # code --install-extension --force ritwickdey.live-sass
    # code --install-extension --force PKief.material-icon-theme
    # How to manually install VS Code extension packaged in a .vsix file:
    # install using the VS Code command line providing the path to the .vsix file.
    # code --install-extension myExtensionFolder\myExtension.vsix
    # --force will force an extension to update if already installed
    # - Languages: AutoHotkey, Ini for VSCode, LaTeX Workshop, PowerShell, Python, REG, SQL Server (mssql), "
    # - Beautify, Rainbow Brackets, Rainbow Highlighter"
    # - Rainglow (dozens of themes), Rainbow Theme"

    # Install the latest .NET Core SDK to build .NET apps online
    choco install dotnetcore-sdk

    # Welcome to .NET Core 3.1!
    # ---------------------
    # SDK Version: 3.1.403
    # 
    # Telemetry
    # ---------
    # The .NET Core tools collect usage data in order to help us improve your experience. The data is anonymous. It is collected by Microsoft and shared with the community. You can opt-out of telemetry by setting the DOTNET_CLI_TELEMETRY_OPTOUT environment variable to '1' or 'true' using your favorite shell.
    # 
    # Read more about .NET Core CLI Tools telemetry: https://aka.ms/dotnet-cli-telemetry
    # 
    # ----------------
    # Explore documentation: https://aka.ms/dotnet-docs
    # Report issues and find source on GitHub: https://github.com/dotnet/core
    # Find out what's new: https://aka.ms/dotnet-whats-new
    # Learn about the installed HTTPS developer cert: https://aka.ms/aspnet-core-https
    # Use 'dotnet --help' to see available commands or visit: https://aka.ms/dotnet-cli-docs
    # Write your first app: https://aka.ms/first-net-core-app
    # --------------------------------------------------------------------------------------
    # Could not execute because the specified command or file was not found.
    # Possible reasons for this include:
    #   * You misspelled a built-in dotnet command.
    #   * You intended to execute a .NET Core program, but dotnet-p-default does not exist.
    #   * You intended to run a global tool, but a dotnet-prefixed executable with this name could not be found on the PATH.
}

function Install-7Zip {

    # https://silentinstallhq.com/7-zip-19-00-silent-install-how-to-guide/
    # For silent uninstall, /S works, but /s does not
    # Possibly use PeaZip instead of 7-Zip?
    if ((New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator) -ne $true) {
        "Chocolatey operations require Administrator elevation.`nTry 'sudo <command>   , or   sudo { <command1> ; <command2> ; ... }`n"
        break
    }

    # Note: if you use "ps 7zFM" this will generate an error, but "ps *7zFM*" will not as it is a wildcard sweep!
    while (ps *7zFM*) { read-host "Cannot contiune while 7-Zip Manager is open, please close manually then continue, or Ctrl-C to quit this function." }

    "7-Zip is closed so setup will now continue (do not reopen 7-Zip until setup completes)."

    # Test if choco package is installed, if not, then it must have been installed manually, perform uninstall
    $result = C:\ProgramData\Chocolatey\choco list -lo | Where-object { $_.ToLower().StartsWith("7Zip".ToLower()) }
    if ($result -ne $null) {
        # Check if C:\Program Files\7-Zip\7zFM.exe exists, if not, need to uninstall the choco package then reinstall
        if (!(Test-Path "C:\Program Files\7-Zip\7zFM.exe")) {
            choco uninstall -y 7zip --force
            choco install -y 7zip --force
        }
        else {
            choco upgrade -y 7zip   # Note that 'update' is a deprecated term, always use 'upgrade'
        }
    }
    else
    {
        "No chocolatey 7-Zip package was found, will remove versions if present then reinstall`n"
        # Remember, must use /S and not /s for silent uninstall.
        $command = @'
cmd.exe /C "C:\Program Files (x86)\7-Zip\uninstall.exe" /S
'@ ; if (Test-Path "C:\Program Files (x86)\7-Zip\uninstall.exe") { Invoke-Expression -Command:$command }
        
        $command = @'
cmd.exe /C "C:\Program Files\7-Zip\uninstall.exe" /S
'@ ; if (Test-Path "C:\Program Files\7-Zip\uninstall.exe") { Invoke-Expression -Command:$command }

        # Now install the latest chocolatey package. Note: installs to C:\Program Files. i.e. not to the "(x86)" folder.
        choco install -y 7zip --force
    }
}
Set-Alias Install-7-Zip Install-7Zip
Set-Alias Install-7z Install-7Zip

function Install-7ZipDownloadOnly {
    # Three templates to download from a website.
    # Using 7-Zip download website as an example.
    # Keep this for reference and probably for VS Code Portable, BitComet Portable, etc
    $start_time = Get-Date   # Used with the timer on last line of function

    # https://www.reddit.com/r/PowerShell/comments/9gwbed/scrape_7zip_website_for_the_latest_version/
    # Modified each version to use common $url / $page variables:
    $url = "https://www.7-zip.org/download.html"
    $page = Invoke-WebRequest -uri $url

    ### Method 1: Parsing the HTML
    $table  = $page.ParsedHtml.getElementsByTagName('table')[0]
    $ver    = $table.cells[1].getElementsByTagName('p')[0].innertext.split()[2,3]
    "Current Version: {0} {1}" -f $ver[0], $ver[1]
    
    # Method 2: Use regex ( -match ) with $page
    # Note on regex with html:
    # https://stackoverflow.com/questions/1732348/regex-match-open-tags-except-xhtml-self-contained-tags
    $regex  = $page.Content -match 'Download 7-Zip (.*)\s(.*) for Windows'
    if ($regex) {
        $ver  = $Matches[1]
        $date = $Matches[2]
    }
    "Current Version: {0} {1}" -f $ver, $date

    ### Method 3: Get a full download link
    $urlPrefix = ($url -Split('download'))[0]   # Get everything before download in the link
    $RelativeLink = (((($page -Split '7-Zip for 64-bit Windows'   # Useful technique to drill down to what we require
        )[0] -split '<TR>'
        )[-1] -split 'href="'
        )[1] -split '"'
        )[0]
    
    $DL_ver = $RelativeLink.Split('z')[1].Split('.')[0]
    $DL_link = -Join($urlPrefix, $RelativeLink)
    $DL_file = ($DL_link -Split('/'))[-1]
    
    $DL_ver
    $DL_link
    $start_time = Get-Date
    iwr -Uri $DL_link -OutFile .\$DL_file

    Write-Output "Time to download: $((Get-Date).Subtract($start_time).Seconds) second(s)"   # Compact timer technique for a function
}

function Install-BitCometPortable {
    $start_time = Get-Date   # Used with the timer on last line of function
    $url = "https://www.bitcomet.com/en/archive"
    # Solving the First-Launch Configuration Error with PowerShell Invoke-WebRequest Cmdlet
    # https://stackoverflow.com/questions/38005341/the-response-content-cannot-be-parsed-because-the-internet-explorer-engine-is-no
    # https://wahlnetwork.com/2015/11/17/solving-the-first-launch-configuration-error-with-powershells-invoke-webrequest-cmdlet/
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 2
    $page = Invoke-WebRequest -uri $url -UseBasicParsing   # Without the above change, function will fail if IE has never run

    ($page.Content -match 'https\:\/\/download\.bitcomet\.com\/achive\/BitComet_\d\.\d\d\.zip')
    $dl_link = $Matches[0]
    $dl_ver = (($dl_link.Split('_')[1]) -split '.zip')[0]
    $dl_file = ($dl_link -Split('/'))[-1]
    $dl_out = ($dl_file -split "_")[0]
    $dl
    $dl_ver
    $dl_file
    iwr -Uri $dl_link -OutFile .\$dl_file
    # Maybe test here: f not exist 7z.exe run Install-7zip to make sure it is available
    $sevenzip = 'C:\Program Files\7-Zip\7z.exe'
    & $sevenzip "x" $dl_file "-o./$dl_out" "-y" "-r"
    & "$dl_out\BitComet.exe"
    Write-Output "Time to configure: $((Get-Date).Subtract($start_time).Seconds) second(s)"
}

function AutoScrapeMagnetLinks {
    # https://stackoverflow.com/questions/19368028/downloading-a-magnetic-link-with-powershell-3
    # start "magnet:?xt=urn:btih:44bb5e0325b7dad0bdc5abce459b85b014766ec0&dn=MY_TORRENT&tr=udp%3A%2F%2Ftracker.openbittorrent.com%3A80&tr=udp"
    # https://sonarr.tv/#home
    # https://github.com/Sonarr/Sonarr
    # https://github.com/jpmikkers/FirefoxMagnetMimeHandler
    # https://www.reddit.com/r/sonarr/comments/f1l7g3/update_windows_powershell_script_for_qbittorrent/
}

function Install-WinSCP {
    <#
    .SYNOPSIS
    #>
}

function Install-Chocolatey {
    if (Test-Path C:\ProgramData\chocolatey\bin\choco.exe) {
        "Chocolatey is installed at C:\ProgramData\chocolatey"
    }
    else {
        # Below is long form, works on PowerShell 1+
        # For PowerShell 3+, can use: iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))"
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }
}

function Help-JustInstallCore {
@"
Very lightweight alternative to chocolatey with much smaller selection of tools but all very useful.

msiexec.exe /i https://just-install.github.io/stable/just-install.msi

https://just-install.github.io/
"@ | more
}

function Help-ScoopCore {
    # IfExistSkipCommand "$($env:USERPROFILE)\scoop\apps\scoop\current\bin\scoop.ps1" "Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')"
    # if (Test-Path("$($env:USERPROFILE)\scoop\apps\scoop\current\bin\scoop.ps1")) {
    @"
scoop https://scoop.sh/ is a PowerShell based package manager.
Set-ExecutionPolicy RemoteSigned -scope CurrentUser

Long Install  : Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
Short Install : iwr -useb get.scoop.sh | iex   # Requires PS 3+

scoop is more console-app-focussed than chocolatey (but various GUI apps are available)

Docs: https://github.com/lukesampson/scoop/wiki
Apps: https://github.com/ScoopInstaller/Main/tree/master/bucket

The main scoop app is at: $($env:USERPROFILE)\scoop\apps\scoop\current\bin\scoop.ps1
Packages install to: $($env:USERPROFILE)\scoop\apps\
'scoop help', 'scoop search sql', 'scoop install pshazz' (git-enabled PowerShell prompt tool)
Review: https://dev.to/lgraziani2712/how-scoop-made-my-dev-experience-in-windows-so-great-3k20

To uninstall Scoop and all programs you've installed with Scoop:
   scoop uninstall scoop
If you're sure, just type 'y' and press enter to confirm.
Broken Install: If you delete ~/scoop you should be able to reinstall:
   del ~\scoop -Force  (optionally with -Recurse)
   ~/appdata/local/scoop
concfg notes: https://stackoverflow.com/questions/13690223/how-can-i-launch-powershell-exe-with-the-default-colours-from-the-powershell-s

Sample advanced install of 'Zeal':
First you need to enable Scoop's extras bucket, if that wasn't done before:
   scoop bucket add extras
To install Zeal run the following command:
   scoop install zeal
Scoop can also install Visual C++ 2015 Redistributable:
   scoop install vcredist2015
"@ | more
    # Pshazz was mentioned here.
    # https://stackoverflow.com/questions/5725888/windows-powershell-changing-the-command-prompt
    # https://github.com/JanDeDobbeleer/oh-my-posh
    # https://github.com/lukesampson/pshazz
}

function Help-ChocoCore {
<#
.SYNOPSIS
Outputs information on Chocolatey apps
#>
@"

The Following one-liner will setup Chocolatey:
   Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

   choco list --local-only       # List all installed packages    # choco list -lo
   choco upgrade all --noop      # List all packages, versions, and available upgrades (note, this command replaces the older 'choco version all' command)

It is often useful to schedule a daily job to upgrade all packages 'choco upgrade all -y' daily during the night:

Multiple installs can be shortened to a single line:
    cinst notepadplusplus sysinternals java firefox chrome -y

Most application packages will install to their default install location (i.e. "C:\Program Files (x86)\Notepad++" for the notepadplusplus package)

However, packages that are standalone tools (SysInternals, gsudo, etc) that have no specific install location use a 'shim' per executable.
A package will be in a single folder under. e.g. 'C:\ProgramData\Chocolatey\lib\gsudo'.
Binaries are then placed into a 'tools' subfolder under there. e.g. 'C:\ProgramData\Chocolatey\lib\gsudo\tools'.
Then, a 'shim' is created for each binary under 'C:\ProgramData\Chocolatey\bin'.
Only the bin folder containing the various shim executables are is added to the PATH statement.
Some apps don't do this though. e.g. gsudi actually creates its binary in 'C:\ProgramData\Chocolatey\lib\gsudo\bin'.
And that is on the PATH statement (which seems inefficient but usually they only do this if shims are not possible).

   cinst -y chocolatey.gui      # GUI tool to find packages
   cinst -y chocolatey.server   # Create a standalone Chocolatey server on a PC
   cinst -y boxstarter          # Boxstarter allows automating installs including through reboots

https://github.com/chocolatey/chocolatey.org   # Central location for the Chocolatey Project (package source etc)
https://github.com/chocolatey/choco            # github source
https://github.com/chocolatey/choco/issues     # github issues (post any bugs or feature requests here)
https://gitter.im/chocolatey/home              # Various Chocolatey/Boxstarter chat forums on Gitter
https://gitter.im/chocolatey/choco             # Gitter for choco command line
https://gitter.im/chocolatey/chocolatey.org    # Gitter for chocolatey.org site
https://chocolatey.org/profiles/bcurran3       # Prolific package authors list of packages
https://us8.list-manage.com/subscribe?u=86a6d80146a0da7f2223712e4&id=73b018498d   # Mailing list

Selected packages:
   cinst -y 7zip notepadplusplus GoogleChrome Firefox microsoft-windows-terminal bitdefenderavfree greenshot
   cinst -y vscode visualstudiocode javaruntime jre8 autohotkey python3 
   cinst -y anydesk teamviewer citrix-receiver hamachi zerotier-one filezilla filezilla.server skype
   cinst -y synergy ultramon displayfusion inputdirector sharemouse
   cinst -y mpc-hc-clsid2 bbc-iplayer handbrake ffmpeg virtualdub yacreader Paint.NET Gimp ImageMagick
   cinst -y gsudo nircmd kindlegen calibre regjump produkey performancetest hwinfo
   cinst -y shutup10 recuva wox rainmeter    # Shutup (control Windows telemetry data passed to Microsoft), recuva (Hard disk recovery app), wox (very useful launcher app), rainmeter (customise desktop)  # https://www.slant.co/versus/5390/11687/~wox_vs_launchy
   cinst -y PowerToys vlc rufus   # Win10 PowerToys, VLC (I prefer MPC-HC), Rufus (create bootable USB stick)
   cinst -y git azure-cli azurepowershell nextcloud putty putty.portable
   cinst -y sql-server-express sql-server-management-studio
   cinst -y DotNet4.5.2 DotNet4.6.1 dotnet4.6.2 KB2919355 KB2919442 KB2999226 KB3033929 KB3035131 
   cinst -y vcredist140 vcredist2013 vcredist2015 visualstudio2017-installer visualstudio2017community
   cinst -y steam -ia "INSTALLDIR=""D:\0 Cloud\Steam"""            # Install Steam to my preferred custom folder
   cinst -y goggalaxy -ia "INSTALLDIR=""D:\0 Cloud\GOG Galaxy"""   # Install GOG Galaxy to my preferred custom folder
   cinst -y windows-sandbox vmwareplayer                           # Virtualisation
   cinst -y Microsoft-Hyper-V-All -source windowsFeatures          # Hyper-V using WindowsFeatures as a source
   cinst -y wsl               # Alternatively:    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux
   cinst -y wsl-ubuntu-1804 wsl-alpine wsltty   # Download Linux distros for WSL

   cinst -y dotnet3.5 dotnet4.5 vcredist2005 vcredist2008 vcredist2010 vcredist2012 vcredist2013   # dotnet4.6 is redundant, part of Win10 build
   cinst -y rdmfree terminals rdcman rdtabs mRemoteNG tightvnc ultravnc filezilla filezilla.server
   cinst -y lightshot greenshot sharex screenpresso
   cinst -y chromium 
   cinst -y wamp-server tomcat apache
"@ | more
}

function Help-ChocoDB {
@"

mssqlserver2014-sqlocaldb   # LocalDB is a lightweight version of Express, all programmability features yet runs in user mode with fast zero-config install. Use this if you need a simple way to create and work with DBs from code
MsSqlServer2014Express      # Sql Server
mysql   # choco install mysql --params "/port:3307 /serviceName:AltSQL"
mariadb # drop-in replacement of MySQL with more features, new storage engines, fewer bugs, and better performance. MariaDB server is a community developed fork of MySQL server started by core members of the original MySQL team.
mongodb # choco install packageID [other options] --params="'/ITEM:value /ITEM2:value2 /FLAG_BOOLEAN'"). To have choco remember parameters on upgrade, be sure to set choco feature enable -n=useRememberedArgumentsForUpgrades.

"@ | more
}

function Help-ChocoToDo! {
@"

Sorting, work in progress ...
Find more Azure, AWS, git, armclient (Azeure Resource Manager API)
Find more puppet - ....
Find more apache-httpd
Find more dvd cdex (audio CD rip)
"@ | more
}

function Help-ChocoSysTools {
@"

Probably recommend not to install these via Chocolatey.
e.g. wsl package is probably best installed using PowerShell directly.
wsl                 # Widnows Subsystem for Linux

cinst -y shutup10   # tool to manage all Windows telemetry and what data is passed to Microsoft
cinst -y recuva.portable   # Hard disk recovery app (when folders accidentally overwritten etc)
cinst -y wox        # Wox seems about the best of the launcher apps, huge amoutn of features and extensibility.
    # https://www.slant.co/versus/5390/11687/~wox_vs_launchy
cinst -y rainmeter  # Customise and skin Windows
cinst -y rktools.2003   # Resource Kit Tools 2003
cinst -y PowerToys      # The new PowerToys for Windows 10 0.11.0
"@ | more
}

function Help-ChocoFileSearch {
@"

astrogrep    # regex file search GUI 
ag           # regex file search console "silversearcher"
ant-renamer  # multiple file renamer GUI
advanced-renamer

"@ | more
}

function Help-ChocoGames {
@"

https://chocolatey.org/packages?q=games
https://steamcommunity.com/sharedfiles/filedetails/?id=1167945
https://developer.valvesoftware.com/wiki/Command_Line_Options

### Note on organisation: To keep my OS clean / easy to rebuild, I install games outside of the C: drive, so I put 
### Steam/Uplay/Gog into a "cloud apps" folder (as they download and sync from online), so I put these into
### "D:\0 Cloud" (the "0" is just to sort some folders before other folders, similarly "D:\0 Backup"). With this
### setup, I can either install Steam to "C:\Program Files\Steam" and then move "steamapps" to "D:\0 Cloud\Steam\steamapps",
### or, even easier, I just use Chocolatey to install the whole application to "D:\0 Cloud\Steam". In doing this, no
### game files are affected, and I can rebuild the C: drive at any time then just reinstall the app, and the games will
### work as they did previously (so I can keep these apps *out* of my core Windows build and easily reinstall when wanted).

### Multi-Game Launchers
# cinst -y steam
cinst -y steam -ia "INSTALLDIR=""D:\0 Cloud\Steam"""            # Install Steam to my preferred custom folder (tested and working)
# cinst -y uplay
cinst -y uplay -ia "INSTALLDIR=""D:\0 Cloud\Uplay"""            # Install Ubisoft Uplay to my preferred custom folder (to test/fix)
cinst -y goggalaxy          
cinst -y goggalaxy -ia "INSTALLDIR=""D:\0 Cloud\GOG Galaxy"""   # Install GOG Galaxy to my preferred custom folder (to test/fix)
cinst -y epicgameslauncher  # The Epic Games Launcher obtains the Unreal Game Engine, modding tools and other Epic Games like Fortnite and the new Epic Games Store (Epic Games Account Required).
cinst -y legendary          # Alternative Epic Games Launcher        
cinst -y bethesdanet        # Bethesda Games
cinst -y battle.net         # Login for World of Warcraft, StarCraft II, Diablo III, and Hearthstone: Heroes of Warcraft
cinst -y itch               # Install, update and play indie games.
cinst -y butler             # butler is the itch.io command-line tools for uploading games to itch.io
cinst -y cockatrice         # Multiplatform supported program for playing tabletop card games over a network in C++/Qt with support for both Qt4 and Qt5.
cinst -y gdlauncher         # Minecraft launcher. 
cinst -y voobly             # Voobly gaming network launcher, to play various games like Age of Empires II with other people.
cinst -y playnite           # Game library manager and launcher with support for Steam, GOG, Origin, Battle.net and Uplay.
cinst -y minetest           # Open source voxel game engine. Play or mod a game to your own liking, and play on multiplayer server.

### Gaming Tools
cinst -y twitch                       # Twitch gaming community
cinst -y steamcmd
cinst -y steamlibrarymanager.portable # Open source app to manage Steam, Origin and Uplay libraries.
cinst -y depressurizer                # Manage large Steam game libraries with supports manual entry of games from other platforms, so you can categorize, filter and launch all of your (Steam, Origin, uPlay, GOG, etc) games from one location.
cinst -y steam-cleaner                # PC utility for restoring disk space from various game clients like Origin, Steam, Uplay, Battle.net, GoG and Nexon.
cinst -y borderlessgaming             # Borderless Gaming Monitor app
cinst -y gbm                          # Game Backup Monitor. Automatically backup your saved games with optional cloud support.
cinst -y gamesavemanager              # GameSave Manager, you can easily backup, restore and transfer your gamesave(s)
cinst -y ludusavi                     # Backup tool for PC game saves.
cinst -y gamesnostalgia.extension     # This extension primary function is to return download string from gamesnostalgia.com.
cinst -y ds4windows                   # configure a game controller ready for games
cinst -y elgato-game-capture          # Record and Stream from Elgato capture cards.
cinst -y gamebooster                  # Optimize your PC for smoother, more responsive game play.
cinst -y game-collector               # Game Database, catalog your game collection.
cinst -y gamedownloader.install       # Portable and open source download client. # gamedownloader.portable
cinst -y game-key-revealer.portable   # Recover Game Keys
cinst -y gameplay-time-tracker        # Learn time spent on computer games.
cinst -y gamesavemanager              # GameSave Manager, backup, restore and transfer your gamesave(s). 
cinst -y geforce-game-ready-driver    # NVIDIA GeForce Game Ready Driver
cinst -y cheatengine                  # Open source tool designed to help modify single player games to make harder or easier (e.g: Find that 100hp is too easy, try playing a game with a max of 10 HP), and also contains other useful tools to debug games and even normal applications.
cinst -y antimicro                    # Graphically map keyboard buttons and mouse controls to a gamepad.   # antimicro.install, antimicro.portable
cinst -y darksoulsmapviewer           # Dark Souls Map Viewer with DLC content and online map analytics.
cinst -y display-changer              # Changes the display resolution, runs a program, then restores the original settings.
cinst -y gbstudio                     # Retro adventure game creator for handheld video game system.   # gbstudio.portable
cinst -y gdevelop                     # Open source game development, create HTML5 and native games.
cinst -y project-aurora               # Unified lighting effects for Logitech, Corsair, Razer, Clevo, Cooler Master, Steelseries, Wooting, Roccat, Alienware, PlayStation 4, Drevo, Soundblaster X, Asus, Yeelight and NZXT ecosystems.
cinst -y uniws                        # Universal widescreen patcher patches games to properly support widescreen resolutions
cinst -y voicebot                     # Voice Powered Game Control
cinst -y ds4windows                   # DS4Windows is a portable app for using DualShock 4 on PC. By emulating a Xbox 360 controller, many more games are accessible.
cinst -y wowcrypt                     # World of Warcraft Community API Project. :: World of Warcraft Armory Desktop Application
cinst -y wow-stat                     # World of Warcraft server uptime monitor.
cinst -y geforce-experience           # The easiest way to update your drivers, optimise your games, and share your victories.
cinst -y nvidia-geforce-now           # Web based game streaming, powered by NVIDIA.
cinst -y nvidia-profile-inspector     # modifying game profiles inside the internal driver database of the nvidia driver. 
cinst -y opentrack                    # head tracking software to relay information to games and flight simulation software.
cinst -y openal                       # Cross-platform 3D audio API appropriate for use with gaming applications and many other types of audio applications.
cinst -y moddb.extension              # moddb.com website functions to extract download url from moddb.com website for a specific Mod.
cinst -y kill-frozen-programs         # Easily terminate frozen programs with gamers in mind for highest-priority always-on-top games.
cinst -y mo2                          # Mod Organizer (MO) is a tool for managing installation and changing mod collections.
cinst -y godot                        # Multi-platform 2D and 3D game engine.   # Also: godot-mono
cinst -y poi                          # POI is a scalable KanColle browser and tool based on Electron. Basic functionalities to enhance the gaming experience and is complemented by a variety of plugins (behaves the same as Chrome and does not modify game data, packets or implement bots/macros).
cinst -y opera-gx                     # Opera web browser. Opera GX version provides an integrated CPU, RAM and traffic limiter.

### Emulators
cinst -y emulationstation  # Cross-platform graphical front-end for emulators with controller navigation   # emulationstation.portable
cinst -y mame              # MAME (Multi Arcade Machine Emulator)
cinst -y consoleclassix    # Atari, Nintendo, and Coleco games with our software for free
cinst -y DOSbox      # DOS emulator for playing old games
cinst -y DOSbox-x    # DOSbox-X fork
cinst -y dbgl        # DOSBox Game Launcher
cinst -y launchbox   # DOSbox Front End, but has since expanded to support both emulated games and other modern PC games also
cinst -y ScummVM     # Old-School games
cinst -y moboplay    # All-in-One Android & iOS Manager
cinst -y pcsx2       # Playstation 2 Emulator   # Also: cinst -y pcsx2.portable
cinst -y no-cash-psx # Tiny PSone emulator with low system requiremments
cinst -y yabause     # SEGA Saturn Emulator
cinst -y genymotion  # Android Emulator
cinst -y winvice     # C64 Emulator
cinst -y fs-uae      # Amiga Emulator
cinst -y project64   # Nintendo 64 Emulator
cinst -y mupen64plus # Nintendo 64 Emulator
cinst -y bsnes       # SNES Emulator
cinst -y zSNES       # SNES Emulator
cinst -y hakchi2.portable  # This is a GUI for hakchi (a ROM-management tool for NES Mini) by madmonkey.
cinst -y mupen64plu  # Mupen64Plus-Qt utilizes the console UI to launch games. It contains support for most of the command line parameters. These can be viewed by running Mupen64Plus from a terminal with the --help parameter
cinst -y openra      # Free, real-time strategy game engine supporting early Westwood classics, such as Command & Conquer titles.

### Games
cinst -y gzdoom          # Enhanced port of the Doom engine (ZDoom is a family of enhanced ports of the Doom engine)
cinst -y doomsday        # Enhanced Doom/Heretic/Hexen port.
cinst -y qc-doom-edition # Doom Edition mod that brings character classes, and weapons from the latest Quake game into Doom.
cinst -y slade           # SLADE3 isn editor for Doom-engine based games.
cinst -y doom-d64rtr     # Doom for Nintendo 64 ported to Windows.
cinst -y angry-birds-star-wars clash-royale clash-of-clans
cinst -y tetr-io         # Tetris clone.
cinst -y aliengame       # Learn about solving complex problems by playing a game.
cinst -y sil             # Sil is a computer role-playing game with a strong emphasis on discovery and tactical combat.
cinst -y triplea         # TripleA strategy game.
cinst -y azpazeta        # Azpazeta is the strategy-economic game where you can play thousands of adventures with the custom maps.
cinst -y FreeOrion       # FreeOrion is a tribute to the old Master of Orion games
cinst -y angband         # Angband is a free, single-player dungeon exploration game.
cinst -y nethack         # NetHack is a single player dungeon exploration game.
cinst -y redrogue        # Side-scrolling roguelike-like, descend into the Dunngeon of Chaos, retrieve the Amulet of Yendor, kill all who stand in your way and return with your prize.
cinst -y dwarf-fortress -ia "INSTALLDIR=""D:\Dwaarf"""  # Single-player fantasy rogue-like. Control a dwarven outpost or an adventurer in a randomly generated, persistent world.
cinst -y cavesofqud      # Caves of Qud is a far-future roguelike in the tradition of the pen and paper classic, Gamma World.
cinst -y crawl           # Dungeon Crawl Stone Soup is a free roguelike game of exploration and treasure-hunting in dungeons filled with dangerous and unfriendly monsters in a quest for the mystifyingly fabulous Orb of Zot.
cinst -y galaxian        # Galaxian clone.
cinst -y skifree         # Skiing game, avoid obstacles.
cinst -y mrboom          # Mr.Boom is an up to 8 players Bomberman clone.
cinst -y eduactiv8       # eduActiv8 (formerly pySioGame) is a free cross-platform Open Source educational program for children.
cinst -y blitz burgertime centipede donkeykong frogger galaxian lunarlander millipede minipacman monkeyacademy misslecommand pacman snake superburgertime   # PJ Crossley's remakes of the classic games (2017). http://pjsfreeware.synthasite.com
cinst -y bunnyhop        # PJ Crossley original game.
cinst -y sgt-puzzles     # Simon Tatham, collection of small one-player puzzle games
   # Black Box, Bridges, Cube, Dominosa, Fifteen, Filling, Flip, Flood, Galaxies, Guess, Inertia, Keen, Light Up, Loopy, Magnets, 
   # Map, Mines,Net, Netslide, Palisade, Pattern, Pearl, Pegs, Puzzles Manual, Puzzles Web Site, Range, Rectangles, Same Game,
   # Signpost, Singles, Sixteen, Slant, Solo, Tents, Towers, Tracks, Twiddle, Undead, Unequal, Unruly, Untangle
"@ | more
}

function Help-ChocoProgrammingLanguages {
@"

cinst -y autohotkey python3 
cinst -y activeperl strawberryperl 
cinst -y python python2 python3 pygtk-all-in-one_win32_py2.7
cinst -y ruby rails
cinst -y golang
cinst -y clojure.clr dmd scala erland ocaml php5-dev gtksharp qt-sdk-windows-x86-msvc2013_opengl
cinst -y java
cinst -y javaruntime adobeair silverlight 
cinst -y sharpdevelop linqpad   # LINQ

# Note that python3 installs to C:\Python38 and javaruntime installs (both 32-bit and 64-bit versions).
# choco uninstall python -x`. The `-x` mean remove dependencies.
# choco install python3 --installargs='TargetDir=""C:\Program Files\Python3""'
# https://docs.python.org/3/using/windows.html#installing-without-ui
"@ | more
}

function Help-ChocoConsoleTools {
@"

Various console tools

cinst -y gsudo       # Very clean/simple sudo tool for Windows concoles (DOS/PowerShell)
cinst -y nircmd      # NirSoft command line toolkit
cinst -y kindlegen   # Amazon Kindle epub-to-mobi conversion tool
cinst -y regjump     # regjump <registry_location> will open regedit.exe and jumpt to that location
cinst -y regeditor
cinst -y regscanner
cinst -y regfromapp
cinst -y regjump
cinst -y regalyzer
cinst -y rcp-registry
cinst -y advancedrun
cinst -y winmd
cinst -y terminal-icons.powershell   # 2019-06-16 (P) :: A PowerShell module to show file and folder icons in the terminal. :: Terminal-Icons is a PowerShell module that adds file and folder icons when displaying items in the terminal. This relies on the custom fonts provided by [Nerd Fonts](https://github.com/ryanoasis/nerd-fonts).
"@ | more
}

function Help-ChocoRegistry {
@"

Various registry apps

cinst -y regjump       # regjump <registry_location> will open regedit.exe and jumpt to that location
cinst -y regeditor
cinst -y regscanner
cinst -y regfromapp
cinst -y regalyzer
cinst -y rcp-registry
cinst -y advancedrun
cinst -y winmd
"@ | more
}

function Help-ChocoMicrosoftFrameworksAndTools {
@"

cinst -y dotnetcore-sdk # Install the latest .NET Core SDK to build .NET apps online
cinst -y dotnet3.5      # Equivalent DISM command (note that this covers .NET 2.0 and 3.5):  dism /online /Enable-Feature /Featurename:NetFx3
cinst -y dotnet4.5      # To get latest 4.5.x install this package (will install current latest 4.5.2)
cinst -y dotnet4.6      # Not required on Win 10 as part of build
cinst -y vcredist2005 
cinst -y vcredist2008 
cinst -y vcredist2010 
cinst -y vcredist2012 
cinst -y vcredist2013 
cinst -y msaccess2010-redist
cinst -y directx        # Still required by many games
cinst -y webpicmd       # Microsoft Web Platform Installer for IIS
cinst -y rktools.2003   # Resource Kit Tools 2003
cinst -y PowerToys      # 0.11.0 The new PowerToys for Windows 10
"@ | more
}

function Help-ChocoTextEditors {
@"

cinst -y notepadplusplus vim
cinst -y spf13-vim vim-dwiw2015 kickassvim nano gedit
cinst -y textpad notepad2 notepad3 sublime sublimetext3.powershellalias jivkok.sublimetext3.packages
cinst -y bowpad focuswriter emacs xyzzy
cinst -y brackets atom emeditor komodo-edit
cinst -y programmersnotepad notcl mery texts jedit
cinst -y bluefish babelpad nimbletext
cinst -y zim    # Multi-page personal wiki book
cinst -y foxe   # XML Editor
"@ | more
}

function Help-ChocoContainersAndVMTools {
@"

cinst -y vmwareplayer
cinst -y vmwarevsphereclient
cinst -y vmware.powercli
cinst -y vmware-tools      # Only install *inside* Vmware VMs to activate vmware-tools.
cinst -y virtualbox
cinst -y vboxheadlesstray
cinst -y vagrant

cinst -y docker-machine
cinst -y docker-toolbox    # Docker Toolbox is for older Mac and Windows systems that do not meet the requirements of Docker for Mac and Docker for Windows. Docker Toolbox is an installer for quick setup and launch of a Docker environment on older Mac and Windows systems that do not meet the requirements of the new [Docker for Mac](https://docs.docker.com/docker-for-mac/) and [Docker for Windows](https://docs.docker.com/docker-for-windows/) apps by creating a VM and installing Docker inside th
# If installing docker on Windows, it *requires* Hyper-V (does not work with VMware).
cinst -y Microsoft-Hyper-V-All -source windowsFeatures   # Note: using 'WindowsFeatures' as a source instead of Chocolatey
cinst -y docker-for-windows --no-progress --fail-on-standard-error
# Also, script for unattended install on Windows Server
# https://github.com/cdaf/windows/blob/master/automation/provisioning/installDocker.ps1 486
# shutdown /r /t 0

### Sandbox environments
cinst -y windows-sandbox
cinst -y sandboxie
"@ | more
}

function Help-ChocoLinux {
@"

Various Linux related tools and WSL images ready to setup

### WSL (Windows Subsystem for Linux)
# cinst -y wsl
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux
# cinst -y wsl-fedoraremix
# cinst -y wsl-debiangnulinux
# cinst -y wsl-alpine
# cinst -y wsl-kalilinux
# cinst -y wsl-archlinux
# cinst -y wsl-ubuntu-1804
# cinst -y wsl-sles       # SUSE Linux Enterprise Server 12 SP2 for Windows Subsystem for Linux (Install) 12.2.0.020181002
# cinst -y wsl-opensuse   # openSUSE 42.2 (Malachite) for Windows Subsystem for Linux (Install) 42.2.0.020181002
# cinst -y wsltty         # Mintty terminal for WSL
# cinst -y microsoft-windows-terminal   # Can act as WSL and SSH terminal as well as for DOS/PowerShell.

### Linux filesystem tools
# cinst -y linux-reader   # Ext 2/3/4, UFS2, HFS and ReiserFS/4 for Windows
# cinst -y Ext2Fsd        # ext3/4 system driver for windows
# cinst -y enfs4win

### The below are mostly redundant if using WSL
cinst -y gow 
cinst -y lili      # Linux Live USB Installer: create USB installers for all Linux distros
cinst -y wubi      # Windows based Ubuntu Installer
cinst -y xming     # X Server for Windows
"@ | more
}

function Help-ChocoAntiVirus {
@"

Various anti-virus apps

### Anti-virus, adware, malware, trojans, rootkits, etc

cinst -y bitdefenderavfree   # Currently leads most top lists as most reliable AV tool, update this if better are there.
cinst -y unchecky            # A tool to prevent malware by unchecking unrelated offers while installing software

### Other anti-virus
cinst -y avgantivirusfree avirafreeantivirus pandafreeantivirus clamwin kav (Kasperskey Anti-Virus) microsoftsecurityessentials eset.nod32

### Junkware cleanup tools:
cinst -y ccleaner ccenhancer adwcleaner malwarebytes combofix jrt superantispyware avastbrowsercleanup systemninja kvrt windowsrepair
# (Kasperskey Virus Removal Tool)
"@ | more
}

function Help-ChocoRemote-VPN-RDP-VNC-SSH-Terminals {
@"

### Remote Desktops
https://www.msftnext.com/best-remote-desktop-connection-manager-apps-for-windows/
cinst -y anydesk
cinst -y teamviewer
cinst -y citrix-receiver
cinst -y rdmfree
cinst -y terminals
cinst -y rdcman
cinst -y rdtabs
cinst -y mRemoteNG
cinst -y royalts
cinst -y radmin
cinst -y radmin.viewer
cinst -y litemanager-server   # remotely connect to a computer

### VNC
cinst -y tightvnc
cinst -y ultravnc   # Note: UltraVNC does not install silently, forces the installation of the server which must be configured upon installation.

### VPN
cinst -y hamachi
cinst -y zerotier-one

### Terminals for local and remote SSH etc connections
cinst -y microsoft-windows-terminal   # 2020-11-20 (P) 2020-11-20 (A) :: Windows Terminal :: ## Terminal & Console Overview
cinst -y putty   # putty.portable superputty kitty
cinst -y fluent-terminal   # 2020-05-27 (P) 2020-05-27 (A) :: Windows console replacement :: A Terminal Emulator based on UWP and web technologies. Note that this is a sideloaded Windows Store App, and Windows 10 is required.
cinst -y FluentAutomation.Repl :: 0.1.0.3 :: Published 2014-02-22 :: FluentAutomation.Repl :: REPL (Read, Eval, Print, Loop) for FluentAutomation, incredibly useful for testing automation commands before building them into tests.
cinst -y poderosa-terminal-net40   # 2016-12-05 (P) :: SSH Client for Windows :: In 2005, adoption of MITOU project (Japanese government choose ambitious project from the public). But after ten years from then, in 2016, it realized complete renewal, major version became 5. Although Poderosa is not open source but proprietary software, you can indefinitely use it as free trial.

"@ | more
}
    
function Help-ChocoScreenCapture {
@"

cinst -y lightshot     # screenshot selected area, upload to server and get short link right away.
cinst -y greenshot     # fast screen capture tool
cinst -y sharex        # fast screen capture tool
cinst -y screenpresso  # Screen capture, including scrolling parts
"@ | more
}
function Help-ChocoKeyboardVideoMouse {
@"
    
# Using a single Keyboard-Mouse with multiple systems
cinst -y synergy
cinst -y ultramon
cinst -y displayfusion
cinst -y inputdirector
cinst -y sharemouse
"@ | more
}

function Help-ChocoBrowsers {
@"
    
cinst -y chrome firefox microsoft-edge-insider
cinst -y chromium
cinst -y opera
cinst -y maxthon
cinst -y tor-browser

cinst -y google-translate-chrome
cinst -y grammarly-chrome
cinst -y adblockpluschrome
cinst -y adblockplusfirefox
cinst -y adblockplusopera
# cinst -y adblockplusie   # not required in Win 10
cinst -y flashplayerplugin
cinst -y flashplayeractivex
cinst -y adobeshockwaveplayer
"@ | more
}

function Help-PSGalleryModules {
@"

SpeculationControl   # By: PowerShellTeam msftsecresponse | 504,232,013 downloads | Last Updated: 15/05/2019 | Latest Version: 1.0.14
This module provides the ability to query the speculation control settings for the system.

AzureRM.profile   # By: azure-sdk | 53,121,414 downloads | Last Updated: 08/02/2019 | Latest Version: 5.8.3
Microsoft Azure PowerShell - Profile credential management cmdlets for Azure Resource Manager

NetworkingDsc   # By: PowerShellTeam gaelcolas dsccommunity | 43,440,865 downloads | Last Updated: 16/10/2020 | Latest Version: 8.2.0
DSC resources for configuring settings related to networking.

PSWindowsUpdate   # By: MichalGajda | 43,097,906 downloads | Last Updated: 20/04/2020 | Latest Version: 2.2.0.2
This module contain cmdlets to manage Windows Update Client.

PackageManagement   # By: PowerShellTeam alerickson NateLehman krishnayalavarthi | 37,677,222 downloads | Last Updated: 24/04/2020 | Latest Version: 1.4.7
PackageManagement (a.k.a. OneGet) is a new way to discover and install software packages from around the web. It is a manager or multiplexor of existing package managers (also called package providers) that unifies Windows package management with a single Windows PowerShell interface. With PackageManagement, you can do the following. - Manage ... More info

Carbon   # By: pshdo webmd-health-services DecoyJoe | 31,690,856 downloads | Last Updated: 18/11/2020 | Latest Version: 2.9.3
Carbon is a PowerShell module for automating the configuration Windows 7, 8, 2008, and 2012 and automation the installation and configuration of Windows applications, websites, and services. It can configure and manage: * Local users and groups * IIS websites, virtual directories, and applications * File system, registry, and certificate pe... More info

PowerShellGet   # By: PowerShellTeam alerickson | 29,616,703 downloads | Last Updated: 05/09/2020 | Latest Version: 3.0.0-beta10
PowerShell module with commands for discovering, installing, updating and publishing the PowerShell artifacts like Modules, DSC Resources, Role Capabilities and Scripts.

Azure.Storage   # By: azure-sdk | 27,036,761 downloads | Last Updated: 09/10/2018 | Latest Version: 4.6.1
Microsoft Azure PowerShell - Storage service cmdlets. Manages blobs, queues, tables and files in Microsoft Azure storage accounts

Az.Accounts   # By: azure-sdk | 24,777,218 downloads | Last Updated: 18/11/2020 | Latest Version: 2.2.1
Microsoft Azure PowerShell - Accounts credential management cmdlets for Azure Resource Manager in Windows PowerShell and PowerShell Core. For more information on account credential management, please visit the following: https://docs.microsoft.com/powershell/azure/authenticate-azureps

DellBIOSProvider   # By: dcpp.dell | 21,429,235 downloads | Last Updated: 27/08/2020 | Latest Version: 2.3.1
The 'Dell Command | PowerShell Provider' provides native configuration capability of Dell Optiplex, Latitude, Precision, XPS Notebook and Venue 11 systems within PowerShell.

ComputerManagementDsc   # By: PowerShellTeam gaelcolas dsccommunity | 20,852,312 downloads | Last Updated: 05/08/2020 | Latest Version: 8.4.1-preview0001
DSC resources for configuration of a Windows computer. These DSC resources allow you to perform computer management tasks, such as renaming the computer, joining a domain and scheduling tasks as well as configuring items such as virtual memory, event logs, time zones and power settings.

AzureRM.KeyVault   # By: azure-sdk | 18,005,400 downloads | Last Updated: 29/08/2018 | Latest Version: 5.2.1
Microsoft Azure PowerShell - KeyVault service cmdlets for Azure Resource Manager

PSLogging   # By: 9to5IT | 17,316,972 downloads | Last Updated: 22/11/2015 | Latest Version: 2.5.2
Creates and manages log files for your scripts.

xCertificate   # By: PowerShellTeam | 16,138,295 downloads | Last Updated: 08/02/2018 | Latest Version: 3.2.0.0
This module includes DSC resources that simplify administration of certificates on a Windows Server

Az.Storage   # By: azure-sdk | 15,514,775 downloads | Last Updated: 20/12/2019 | Latest Version: 4.0.2-preview
Microsoft Azure PowerShell - Storage service data plane and management cmdlets for Azure Resource Manager in Windows PowerShell and PowerShell Core. For more information on Resource Manager, please visit the following: https://docs.microsoft.com/azure/azure-resource-manager/ For more information on Storage, please visit the following: https://docs... More info

xPowerShellExecutionPolicy   # By: PowerShellTeam | 15,175,979 downloads | Last Updated: 25/07/2018 | Latest Version: 3.1.0.0
This DSC resource can change the user preference for the Windows PowerShell execution policy. THIS MODULE HAS BEEN DEPRECATED It will no longer be released. Please use the "PowerShellExecutionPolicy" resource in ComputerManagementDsc instead.

Az.Resources   # By: azure-sdk | 14,902,410 downloads | Last Updated: 20/12/2019 | Latest Version: 4.0.2-preview
Microsoft Azure PowerShell - Azure Resource Manager and Active Directory cmdlets in Windows PowerShell and PowerShell Core. For more information on Resource Manager, please visit the following: https://docs.microsoft.com/azure/azure-resource-manager/ For more information on Active Directory, please visit the following: https://docs.microsoft.com/a... More info

PSDscResources   # By: PowerShellTeam | 13,332,253 downloads | Last Updated: 26/06/2019 | Latest Version: 2.12.0.0
This module contains the standard DSC resources. Because PSDscResources overwrites in-box resources, it is only available for WMF 5.1. Many of the resource updates provided here are also included in the xPSDesiredStateConfiguration module which is still compatible with WMF 4 and WMF 5 (though that module is not supported and may be removed in the ... More info

Az.Automation   # By: azure-sdk | 13,048,609 downloads | Last Updated: 25/08/2020 | Latest Version: 1.4.0
Microsoft Azure PowerShell - Automation service cmdlets for Azure Resource Manager in Windows PowerShell and PowerShell Core. For more information on Automation, please visit the following: https://docs.microsoft.com/azure/automation/

Az.AnalysisServices   # By: azure-sdk | 13,001,193 downloads | Last Updated: 14/07/2020 | Latest Version: 1.1.4
Microsoft Azure PowerShell - Analysis Services cmdlets for Windows PowerShell and PowerShell Core. For more information on Analysis Services, please visit the following: https://docs.microsoft.com/azure/analysis-services/
"@
}

function Help-ChocoSystemAndBenchmarking {
@"
aws-monitor-diskusage # 2016-05-16 Amazon no longer provides their monitoring  scripts for Windows.  Contains an updated AWS disk monitoring script for Windows and schedules it to run every 5 minutes to post metrics to AWS Cloudwatch. :: ATTENTION: This is a licensing compliant, modified, derative work, please see "UPDATE LOG" in the script comments for the changes made.
# https://www.cpubenchmark.net/cpu_list.php   # Enter model number to get comparisons

cinst -y speccy 
cinst -y siw                    # System Information for Windows
cinst -y aida64-business
cinst -y aida64-engineer
cinst -y winscreenfetch         # Cut down version, better to use # import-module screenfetch
cinst -y hardware-freak         # 2016-10-30 Hardware Freak is a free system information utility designed to present you a lot of information about the hardware found inside your PC. :: Hardware Freak is a free system information utility designed to present you a lot of information about the hardware found inside your PC. This tool will give you information about the CPU, BIOS, Motherboard, RAM Memory, Graphics card, Sound card, Hard Drive, External storage devices, Operating system, Optical Media, Networking, Printers, Keyboard, Mouse and USB Ports.
cinst -y hardware-identify      # 2019-04-24 A program to identify and give information about your drivers. :: ___
cinst -y libre-hardware-monitor # 2020-10-09 Fork of Open Hardware Monitor.
cinst -y OpenHardwareMonitor    # 2020-06-19 Open source software that monitors temperature sensors, fan speeds, voltages, load and clock speeds of a computer :: Open Hardware Monitor supports most hardware monitoring chips found on todays mainboards. The CPU temperature can be monitored by reading the core temperature sensors of Intel and AMD processors. The sensors of ATI and Nvidia video cards as well as SMART hard drive temperature can be displayed. The monitored values can be displayed in the main window, in a customizable desktop gadget, or in the system tray. The free Open Hardware Monitor software runs on 32-bit and 64-bit Microsoft Windows XP / Vista / 7 and any x86 based Linux operating systems without installation.
cinst -y belarcadvisor :: 8.5 :: Published 2016-01-18 :: The Belarc Advisor builds a detailed profile of your installed software and hardware, network inventory, missing Microsoft hotfixes, anti-virus status, security benchmarks, and displays the results in your Web browser :: **Belarc**, located in Maynard, MA, allows users to simplify and automate the management of all of their desktops, servers and laptops throughout the world, using a single database and Intranet server. Belarc's products automatically create an accurate and up-to-date central repository (CMDB), consisting of detailed software, hardware and security configurations.
cinst -y cpu-z 
cinst -y dataram-ramdisk       # 2017-03-16 Create a virtual disk drive in RAM
cinst -y winfontsview :: 1.10 :: Published 2015-04-02 :: View samples of Windows fonts installed on your system :: WinFontsView is a small utility that enumerates all fonts installed on your system, and displays them in one simple table.
cinst -y cinebench :: 20.0.0.1 :: Published 2019-10-01 :: Cinebench is a real-world cross platform test suite that evaluates your computer's performance capabilities. :: Cinebench is a real-world cross-platform test suite that evaluates your computer's hardware capabilities. Improvements to Cinebench Release 20 reflect the overall advancements to CPU and rendering technology in recent years, providing a more accurate measurement of Cinema 4D's ability to take advantage of multiple CPU cores and modern processor features available to the average user. Best of all: It's free.
cinst -y heaven-benchmark :: 4.0.0.20200922 :: Published 2020-09-22 :: Extreme performance and stability test for PC hardware: video card, power supply, cooling system. :: Heaven Benchmark is a GPU-intensive benchmark that hammers graphics cards to the limits. This powerful tool can be effectively used to determine the stability of a GPU under extremely stressful conditions, as well as check the cooling system's potential under maximum heat output.
cinst -y SuperBenchmarker :: 4.5.1 :: Published 2018-04-26 :: Latest Approval 2020-07-12 :: Load generator command-line tool for testing websites and HTTP APIs :: Superbenchmarker is a load generator command-line tool for testing websites and HTTP APIs and meant to become Apache Benchmark (ab.exe) on steriod.
cinst -y valley-benchmark :: 1.0.0.20200922 :: Published 2020-09-22 :: Extreme performance and stability test for PC hardware: video card, power supply, cooling system. :: The forest-covered valley surrounded by vast mountains amazes with its scale from a bird's-eye view and is extremely detailed down to every leaf and flower petal.
cinst -y batteryinfoview :: 1.23 :: Published 2017-08-11 :: Latest Approval 2020-07-09 :: View battery information on laptops/netbooks :: BatteryInfoView is a small utility for laptops and netbook computers that displays the current status and information about your battery.
cinst -y bginfo :: 4.28 :: Published 2019-09-23 :: Latest Approval 2020-08-25 :: Generates relevant information about a Windows computer on the desktop's background. :: BgInfo is a fully-configurable program to automatically generate desktop backgrounds that include important information about the system including IP addresses, computer name, network adapters, and more.
cinst -y desktopinfo :: 2.9.0 :: Published 2020-10-05 :: Latest Approval 2020-10-05 :: This little application displays system information on your desktop. :: This little application displays system information on your desktop. Looks like wallpaper but stays resident in memory and updates in real time. Perfect for quick identification and walk-by monitoring of production or test server farms or any computer you're responsible for. Uses very little memory and nearly zero cpu.
cinst -y exeinfo :: 1.01 :: Published 2015-04-03 :: Display general information about executable files :: The ExeInfo utility shows general information about executable files (*.exe), dynamic-link libraries (*.dll), ocx files, and drivers files. It can recognize all major types of executables, including MS-DOS files, New Executable files (16-bit) and Portable Executable files (32-bit).
cinst -y futuremark-systeminfo :: 5.14.693 :: Published 2019-01-03 :: SystemInfo is a component used in many of our benchmarks to identify the hardware in your system. :: SystemInfo is a component used in many of our benchmarks to identify the hardware in your system. It does not collect any personally identifiable information. SystemInfo updates do not affect benchmark scores but you may need the latest version in order to obtain a valid score.
cinst -y HWiNFO32 :: 4.18.1930.1 :: Published 2016-01-08 :: HWiNFO32 - Hardware Information :: ##Deprecated package, install HWiNFO instead.##
cinst -y softwareinformer :: 1.5.1344 :: Published 2020-07-21 :: Get the up-to-date information about the software you actually use :: Software Informer Client (siClient) is a lightweight utility that creates a list of all of your installed programs and their versions, then compares each item on that list against Software Informer's database automatically on a daily basis to identify newer versions of programs, and informs you about any available updates. Please see the [privacy policy](http://software.informer.com/privacy.html) for what information is shared.
cinst -y taskinfo :: 10.0.0.3361678 :: Published 2016-08-07 :: TaskInfo is a poweful utility that combines and improves features of Task Manager and System Information tools :: TaskInfo is a poweful utility that combines and improves features of Task Manager and System Information tools. It visually monitors (in text and graphical forms) different types of system information in any Windows system in real time.
cinst -y hostinfo :: 1.0.1 :: Published 2016-02-04 :: Obtain your computer's internet host name and IP address :: Gathers host information for the local machine: its standard host name (for local access), its network/internet host name, and its IP address.
cinst -y hwinfo :: 6.34 :: Published 2020-11-03 :: Latest Approval 2020-11-03 :: Comprehensive Hardware Analysis, Monitoring and Reporting for Windows and DOS :: In-depth Hardware Information and real time monitoring
aida64-networkaudit :: 6.30.5500 :: Published 2020-10-26 :: Latest Approval 2020-10-26 :: Advanced system diagnostics utility for home users :: Dedicated network audit solution for businesses that supports IT decision-making with essential statistics, and helps companies to reduce their IT costs. With AIDA64 Network Audit, system administrators can make a detailed inventory of the company PC fleet in an automated manner, and track changes in both hardware and software.
auditbeat :: 7.10.0 :: Published 2020-11-12 :: Latest Approval 2020-11-16 :: auditbeat is a lightweight, open source shipper for log file data. :: Contains the chocolatey package for auditbeat
devaudit :: 3.4.0 :: Published 2020-03-23 :: Latest Approval 2020-03-23 :: Identify known vulnerabilities in development packages and applications (NuGet, MSI, Chocolatey, OneGet, Bower)
winaudit :: 3.0.11.0 :: Published 2015-07-20 :: Free inventory and pc audit software :: WinAudit is an inventory utility for Windows computers. It creates a comprehensive report on a machine's configuration, hardware and software. WinAudit is free, open source and can be used or distributed by anyone. IT experts in academia, government, industry as well as security conscious professionals in the armed services, defence contractors, electricity generators and police forces use WinAudit.
https://alternativeto.net/software/speccy/

# Media & Monitor Info
cinst -y mediainfo :: 20.09 :: Published 2020-10-09 :: Latest Approval 2020-10-10 :: MediaInfo supplies technical and tag information about a video or audio file. Supports many audio and video formats, with different methods of viewing information.
cinst -y mediainfo-cli :: 20.09 :: Published 2020-10-09 :: Latest Approval 2020-10-10 :: MediaInfo supplies technical and tag information about a video or audio file. Supports many audio and video formats, with different methods of viewing information.
cinst -y monitorinfoview :: 1.22 :: Published 2020-08-23 :: View essential information about your monitor :: MonitorInfoView is a small utility that displays essential information about your monitor: manufacture week/year, monitor manufacturer, monitor model, supported display modes, and more...
"@ | more
}

function Help-ChocoDisksAndPartitioning {
    @"
    
    # USB image tools
    cinst -y rufus                 # Create bootable USB drives from Windows and Linux images :: Rufus is a utility that helps format and create bootable USB flash drives, such as USB keys/pendrives, memory sticks, etc.
    cinst -y autobootdisk          # 2018-04-06 Bootable USB tool
    
    # Disk Partitioning tools
    cinst -y partitionwizard       # 12.1.01 :: Published 2020-08-28 :: Free partition manager :: MiniTool Partition Wizard Free helps users to manage disks and partitions, check file system, align SSD partition, migrate OS to SSD, clone disk, convert MBR to GPT, etc.
    cinst -y PartitionMasterFree   # 13.5 :: Published 2019-10-21 :: Free (for personal use) and Easy-to-use Disk Management Software. :: EaseUS Partition Master Free Edition is a partition solution and disk management utility. It allows you to extend partition, especially for system drive, settle low disk space problem, manage disk space easily on MBR and GUID partition table (GPT) disk under 32 bit and 64 bit Windows 2000/XP/Vista/Windows 7 SP1/Windows 8. The most popular hard disk management functions are brought together with powerful data protection including: Partition Manager, Disk and Partition Copy Wizard and Partition Recovery Wizard.
    cinst -y diskgenius            # 5.3.0.1066 :: Published 2020-09-24 :: All-in-one solution, partition manager, disk recovery, etc
    
    cinst -y ntfsinfo              # 2016-07-04 Detailed information about NTFS volumes, including the size and location of the Master File Table (MFT) and MFT-zone, as well as the sizes of the NTFS meta-data files. :: NTFSInfo is a little applet that shows you information about NTFS volumes. Its dump includes the size of a drive's allocation units, where key NTFS files are located, and the sizes of the NTFS metadata files on the volume. This information is typically of little more than curiosity value, but NTFSInfo does show some interesting things. For example, you've probably heard about the NTFS equivalent of the FAT file system's File Allocation Table. Its called the Master File Table (MFT), and it is made up of constant sized records that describe the location of all the files and directories on the drive. What's surprising about the MFT is that it is managed as a file, just like any other. NTFSInfo will show you where on the disk (in terms of clusters) the MFT is located and how large it is, in addition to specifying how large the volume's clusters and MFT records are. In order to protect the MFT from fragmentation, NTFS reserves a portion of the disk around the MFT that it will not allocate to other files unless disk space runs low. This area is known as the MFT-Zone and NTFSInfo will tell you where on the disk the MFT-Zone is located and what percentage of the drive is reserved for it.
    
    # Diagnostics
    cinst -y crystaldiskinfo       # 2020-09-28 HDD/SSD utility software
    cinst -y crystaldiskmark       # 2020-06-28 Small HDD/SDD benchmark utility
    cinst -y disk2vhd              # 2.01.0.20160213 :: Published 2016-02-13 :: Disk2vhd simplifies the migration of physical systems into virtual machines (p2v) :: Disk2vhd is a utility that creates VHD (Virtual Hard Disk - Microsoft's Virtual Machine disk format) versions of physical disks for use in Microsoft Virtual PC or Microsoft Hyper-V virtual machines (VMs). The difference between Disk2vhd and other physical-to-virtual tools is that you can run Disk2vhd on a system that's online. Disk2vhd uses Windows' Volume Snapshot capability, introduced in Windows XP, to create consistent point-in-time snapshots of the volumes you want to include in a conversion. You can even have Disk2vhd create the VHDs on local volumes, even ones being converted (though performance is better when the VHD is on a disk different than ones being converted).
    cinst -y diskcountersview      # 1.27 :: Published 2017-01-30 :: Retrieves the S.M.A.R.T information from your IDE/SATA disk :: DiskCountersView displays the system counters of each disk drive in your system, including the total number of read/write operations and the total number of read/write bytes.
    cinst -y diskdump              # 0.1.0 :: Published 2015-06-24 :: Save the raw contents and calculate the CRC32 checksum of each block on a disk drive :: DiskDump is a simple application I created some time ago to assist with dealing with a faulty hard disk. DiskDump not only saves the raw contents of a disk drive, but it also calculates the checksum of each block and stores it in a seperate file. This allows the image to be verified on a block-by-block basis.
    cinst -y diskext               # 1.20 :: Published 2016-07-04 :: Display volume disk-mappings :: DiskExt displays volume disk-mappings by demonstrating the use of the IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS command that returns information about what disks the partitions of a volume are located on (multipartition disks can reside on multiple disks) and where on the disk the partitions are located.
    cinst -y diskmarkstream        # 1.1.2 :: Published 2016-06-24 :: Macro tool for CrystalDiskMark :: DiskMarkStream is an automatic macro tool specially designed for CrystalDiskMark (CDM). It will start CDM with predefined settings and at completion of each test, will save screenshot and log then move on to next test. This consecutive unattended operation will help user to test various conditions and multiple drives at ease.
    cinst -y diskmon               # 2.01 :: Published 2015-12-28 :: This utility captures all hard disk activity or acts like a software disk activity light in your system tray :: This utility captures all hard disk activity or acts like a software disk activity light in your system tray.
    cinst -y disksmartview         # 1.21 :: Published 2016-06-20 :: Retrieves the S.M.A.R.T information from your IDE/SATA disk :: DiskSmartView is a small utility that retrieves the [S.M.A.R.T.](http://en.wikipedia.org/wiki/S.M.A.R.T.) information (S.M.A.R.T = Self-Monitoring, Analysis, and Reporting Technology) from IDE/SATA disks.
    cinst -y diskspd               # 2.0.21 :: Published 2019-05-05 :: DiskSpd is a storage load generator / performance test tool from the Microsoft Windows, Windows Server and Cloud Server Infrastructure Engineering teams. :: DiskSpd is a highly customizable I/O load generator tool that can be used to run storage performance tests against files, partitions, or physical disks. DiskSpd can generate a wide variety of disk request patterns for use in analyzing and diagnosing storage performance issues, without running a full end-to-end workload. You can simulate SQL Server I/O activity or more complex, changing access patterns, returning detailed XML output for use in automated results analysis.
    cinst -y diskview              # 2.40 :: Published 2016-02-10 :: Graphical disk sector utility :: DiskView shows you a graphical map of your disk, allowing you to determine where a file is located or, by clicking on a cluster, seeing which file occupies it. Double-click to get more information about a file to which a cluster is allocated.
    cinst -y disk-wipe             # 1.7 :: Published 2016-11-25 :: Disk Wipe is Free, portable Windows application for permanent volume data destruction. :: ![Screenshot of Disk Wipe](http://www.diskwipe.org/images/diskwipe_screen1.jpg)
    cinst -y dragondisk            # 1.05 :: Published 2013-02-04 :: DragonDisk is a powerful file manager for Amazon S3R, Google Cloud StorageR, and all cloud storage services that provides compatibility with Amazon S3 API.
    cinst -y hpusbdisk             # 2.2.3.20150303 :: Latest Approval 2015-03-02 :: HP USB Disk Storage Format Tool utility will format any USB flash drive, with your choice of FAT, FAT32, or NTFS partition types :: HP USB Disk Storage Format Tool utility will format any USB flash drive, with your choice of FAT, FAT32, or NTFS partition types.
    cinst -y imdisk                # 2.0.10.20181231 :: Latest Approval 2020-07-14 :: ImDisk is a virtual disk driver for Windows. :: ImDisk is a virtual disk driver for Windows NT/2000/XP/Vista/7/8/8.1 or Windows Server 2003/2008/2012. It can create virtual hard disk, floppy or CD/DVD drives using image files or system memory.
    cinst -y ImDisk-Toolkit        # 20.07.27 :: Latest Approval 2020-07-27 :: Requirements: :: Here is the ImDisk Toolkit. This tool will let you mount image files of hard drive, cd-rom or floppy, and create one or several RamDisks with various parameters.
    cinst -y seek-dsc-harddisk     # 1.0.10 :: Latest Approval 2015-07-30 :: :: Custom DSC Resources for managing resources residing on the hard disk
    cinst -y tcp-diskdirextended   # 1.67 :: Published 2019-10-10 :: Total Commander packer plugin that creates a list with all selected files and directories :: Total Commander packer plugin that creates a list file with all selected files and directories, including subdirs - but also lists contents of archive files internally and also lists all the other archive files recognised by installed TC plugins.
    cinst -y testdisk              # 6.14.20131012 :: Published 2013-10-12 :: Data recovery software :: (This package ist deprecated. Use testdisk-photorec instead.) TestDisk is powerful free data recovery software! It was primarily designed to help recover lost partitions and/or make non-booting disks bootable again when these symptoms are caused by faulty software, certain types of viruses or human error (such as accidentally deleting a Partition Table). Partition table recovery using TestDisk is really easy.
    cinst -y testdisk-photorec     # 7.1 :: Published 2019-07-07 :: Latest Approval 2020-08-23 :: Data recovery software :: TestDisk is powerful free data recovery software! It was primarily designed to help recover lost partitions and/or make non-booting disks bootable again when these symptoms are caused by faulty software, certain types of viruses or human error (such as accidentally deleting a Partition Table). Partition table recovery using TestDisk is really easy.
    cinst -y win32diskimager       # 1.0.0.20181220 :: Published 2018-12-20 :: A tool for writing images to USB sticks or SD/CF cards :: This program is designed to write a raw disk image to removable SD or USB flash devices or backup these devices to a raw image file. It is very useful for embedded development, namely Arm development projects (Android, Ubuntu on Arm, etc). Anyone is free to branch and modify this program. Patches are always welcome.
    cinst -y win32diskimager.install  # 1.0.0.20181220 :: Published 2018-12-20 :: A tool for writing images to USB sticks or SD/CF cards :: This program is designed to write a raw disk image to removable SD or USB flash devices or backup these devices to a raw image file. It is very useful for embedded development, namely Arm development projects (Android, Ubuntu on Arm, etc). Anyone is free to branch and modify this program. Patches are always welcome.
    cinst -y win32diskimager.portable # 1.0.0.20181220 :: Published 2018-12-20 :: A tool for writing images to USB sticks or SD/CF cards :: This program is designed to write a raw disk image to removable SD or USB flash devices or backup these devices to a raw image file. It is very useful for embedded development, namely Arm development projects (Android, Ubuntu on Arm, etc). Anyone is free to branch and modify this program. Patches are always welcome.
    cinst -y yandexdisk            # 1.0.0.6 :: Published 2014-11-23 :: Yandex.Disk is a cloud storage created by Yandex. :: Yandex.Disk is a cloud service created by Yandex that lets users store files on "cloud" servers and share them with others online. The service is based on syncing data between different devices.

"@ | more
}

# azure-information-protection-client :: 1.54.59 :: Published 2020-06-10 :: Latest Approval 2020-06-11 :: This package contains the azure information protection client which can be downloaded and installed by everyone. :: The Azure Information Protection Client allows the user to encrypt content within the usual Microsoft Office product family.
    # azure-information-protection-unified-labeling-client :: 2.6.111 :: Published 2020-06-09 :: This package contains the azure information protection unified labeling client which can be downloaded and installed by everyone. :: The Azure Information Protection Client allows the user to encrypt content within the usual Microsoft Office product family.

function Help-ChocoNetworking {
@"        

### TCP/IP, DNS, LAN, PING, WIFI, SSL
cinst -y dns-benchmark :: 1.3.6668.20190425 :: Published 2019-04-29 :: Domain Name Speed Benchmark :: ___
cinst -y lanbench :: 1.1.0.20180725 :: Published 2018-07-25 :: A Simple LAN / TCP Network Benchmark Utility   Features summary: :: ![Screenshot of LANBench](https://web.archive.org/web/20111013083117/http://www.zachsaw.com/images/lanbench.png)
cinst -y informado :: 2.0.0 :: Published 2020-08-08 :: Latest Approval 2020-08-27 :: A tool to read various RSS, Atom and Reddit feeds. :: Use this the tool written in GO to read various RSS feeds. Note that Atom and Reddit feeds can be parsed as well.
cinst -y ipnetinfo :: 1.95 :: Published 2020-08-23 :: Retrieve IP Address Information from WHOIS servers :: IPNetInfo is a small utility that allows you to easily find all available information about an IP address:
cinst -y pinginfoview :: 2.05 :: Published 2020-09-08 :: Ping to multiple host names/IP addresses :: ## pinginfoview
cinst -y tcp-fileinfo :: 2.23 :: Published 2019-10-07 :: Total Commander lister plugin for viewing binary file information :: Version Information viewer plugin for Total Commander.
cinst -y tcp-linkinfo :: 1.5.2 :: Published 2019-10-10 :: Total Commander lister plugin to view and edit *.lnk files :: Total Commander lister plugin to view and edit *.lnk files.
cinst -y tcp-mediainfo :: 1.0.3 :: Published 2019-10-10 :: Total Commander content and lister plugin to retrieve an info from the video and audio files.
cinst -y webcacheimageinfo :: 1.30 :: Published 2019-05-15 :: Latest Approval 2020-08-22 :: Display the software/camera model of images stored in the cache of your Web browser :: WebCacheImageInfo is a simple tool that searches for JPEG images with EXIF information stored inside the cache of your Web browser (Internet Explorer, Firefox, or Chrome), and then it displays the list of all images found in the cache with the interesting information stored in them, like the software that was used to create the image, the camera model that was used to photograph the image, and the date/time that the image was created.
cinst -y wifiinfoview :: 2.65 :: Published 2020-09-15 :: Latest Approval 2020-09-15 :: WiFi Scanner for Windows 7/8/Vista :: WifiInfoView scans the wireless networks in your area and displays extensive information about them, including: Network Name (SSID), MAC Address, PHY Type (802.11g or 802.11n), RSSI, Signal Quality, Frequency, Channel Number, Maximum Speed, Company Name, Router Model and Router Name (Only for routers that provides this information), and more...
cinst -y openssl-wizard :: 1.3 :: Published 2020-08-07 :: A simple GUI to help you with common certificate related tasks :: OpenSSL Wizard is a small gui layer on top of the openssl cli, which allows you to:

### FTP, SSH
cinst -y filezilla filezilla.server
cinst -y winsshd winscp
cinst -y insider

### Team/Chat tools
cinst skype mattermost mumble itch
"@ | more
}

function Help-ChocoWebAndSQLServers {
@"
    
### WAMP, Windows-Apache-MySQL-PHP, Tomcat
cinst -y wamp-server
cinst -y tomcat
cinst -y apache

### IIS
cinst -y webpi     # The Microsoft Web Platform Installer, consolidates all Web Platform components
cinst -y webdeply  # The Web Deployment Tool simplifies the migration, management and deployment of IIS Web servers
cinst -y sqlserver-odbcdriver # Microsoft ODBC Driver 17 for SQL Server is a single dynamic-link library (DLL)

### Microsoft SQL Server
cinst -y sql-server-express     # SQL Server 2019 Express LocalDB lightweight, can be seamlessly upgraded to full SQL
cinst -y sqllocaldb             # SQL Server 2017 Express LocalDB 14.0.1000.169
cinst -y sql-server-2017        # SQL Server 2017 Developer Edition
cinst -y ssrs                   # SQL Server 2017 Reporting Services 
cinst -y ssdt17                 # SQL Server 2017 Data Tools for Visual Studio Update
cinst -y mssqlserver2014express # SQL Server 2014 Express SP3 12.2.5000.20190905
cinst -y sqlserver2014express   # SQL Server 2014 Express, just downloads and starts the installer
cinst -y sqlserverlocaldb       # SQL Server 2012 Express LocalDB 11.0.2318.0
cinst -y sqlsearch              # Free add-in for SMSS to search SQL across DBs.
cinst -y sql-server-management-studio            # SQL Server ???? Management Studio (SSMS) 15.0.18206.0
cinst -y mssqlservermanagementstudio2014express  # SQL Server 2014 Management Studio (SSMS) 12.2.5000.20170905
cinst -y sqlserver2008r2express-managementstudio
cinst -y sql2008r2.nativeclient # Single DLL with SQL OLE DB provider and SQL ODBC driver. 10.53.6560.0
cinst -y sql2012.clrtypes       # SQL Server System CLR Types package 2011.110.3000.1
cinst -y sql2012.powershell     # PowerShell Extensions for SQL Server 2012
cinst -y sql2014-powershell     # PowerShell Extensions for SQL Server 2014

### Open Source DBs
cinst -y mysql      # MySQL Community Edition
    # choco install mysql --params "/port:3307 /serviceName:AltSQL"
cinst -y mariadb    # drop-in replacement of MySQL with more features, new storage engines, fewer bugs, and better performance. MariaDB server is a community developed fork of MySQL server started by core members of the original MySQL team.
cinst -y mongodb    # choco install packageID [other options] --params="'/ITEM:value /ITEM2:value2 /FLAG_BOOLEAN'"). To have choco remember parameters on upgrade, be sure to set choco feature enable -n=useRememberedArgumentsForUpgrades.
cinst -y sqlite     # SQLite is a software library that implements a self-contained, serverless, zero-configuration, transactional SQL database engine. This package also installs sqlite tools by default - sqldiff, sqlite3, sqlite3_analyzer.
cinst -y postgres12 # PostgreSQL
cinst -y postgres   # PostgreSQL
cinst -y firebird   # High performance ANSI SQL standard relational DB for Linux, Windows.

### Tools tied to specific DBs
cinst -y oracle-sql-developer # Free IDE for Oracle DBs.
cinst -y mysql.workbench   # DB Design, Modeling, Development (replacing MySQL Query Browser), Administration (replacing MySQL Administrator)
cinst -y mysql.utilities   # command-line utilities that are used for maintaining and administering MySQL servers.
cinst -y mysql-connector   # develop .NET applications that require secure, high-performance data connectivity with MySQL.
cinst -y mysql-cli         # simpl SQL shell with input line editing.
cinst -y sqlyog            # GUI for MySQL. Community Edition.
cinst -y toad.mysql        # freeware development tool that enables you to rapidly create and execute queries, automate database object management
cinst -y sqlite.shell      # Command-line shell for accessing and modifying SQLite DBs.
cinst -y sqlite.analyzer   # SQLite Analyzer is an analysis program for database files compatible with all SQLite versions through this version and beyond.
cinst -y sqlitebrowser     #  high quality, visual, open source tool to create, design, and edit database files compatible with SQLite.
cinst -y pgadmin3          # Administration and development platform for PostgreSQL.
cinst -y pgadmin4          # Administration and development platform for PostgreSQL.

### Generic DB Tools
cinst -y dbeaver     # Open source and universal DB tool for developers and DBAs.
cinst -y linqpad     # Test any C#/F#/VB snippet or program, query databases in LINQ (or SQL) - SQL/Azure, Oracle, SQLite, Postgres & MySQL
cinst -y heidisql    # Reliable tool designed for web developers using the popular MySQL server, Microsoft SQL databases and PostgreSQL.
cinst -y invantive-data-hub   # command-line driven software that is capable of executing Invantive Query Tool-compatible scripts across many database and cloud platforms.
cinst -y invantive-query-tool # query tool
cinst -y invantive-bridge-connectors-power-bi  # Invantive Bridge Connectors for Power BI
cinst -y databasenet    # intuitive multiple database management tool. With it you can Browse objects, Design tables, Edit rows, Export data and Run queries with a consistent interface.
cinst -y databasenetpro # Trial version of Database .NET Pro.
cinst -y sqltoolbelt    # Redgate's SQL Toolbelt contains the industry-standard products for SQL Server development, deployment, backup, and monitoring. Together, they make you productive, your team agile, and your data safe.
cinst -y dbatools       # PowerShell Module, command-line SQL SMSS, 300 commands.

cinst -y legitest       # LegiTest enables your team to test all aspects of the SQL Server stack. This includes testing of database objects such as stored procedures, functions, views and tables. It further includes all objects on the BI stack, including SSIS Packages, SSAS Cubes and Dimensions (both Tabular and Multidimensional) and SSRS reports.
cinst -y pragmaticworksworkbench       # 2020-07-12 :: Pragmatic Workbench includes BI xPress, DBA xPress, and DOC xPress.  Start with a 14-day free trial or enter your license keys to access this full featured business intelligence suite. :: ### Pragmatic Workbench
cinst -y pragmaticworksworkbenchserver # BI xPress Server, DOC xPress Server, and LegiTest Server are all included in this package, our three powerhouse team focused products.

cinst -y nosql-workbench :: 2.0.0 :: Published 2020-10-29 :: Latest Approval 2020-10-29 :: NoSQL Workbench for Amazon DynamoDB is a cross-platform client-side application for modern database development and operations and is available for Windows and macOS. :: # NoSQL Workbench for Amazon DynamoDB
cinst -y sqlbench :: 1.1.0.0 :: Published 2020-10-18 :: sqlbench measures and compares the execution time of one or more SQL queries. :: #### Overview
cinst -y sql-workbench :: 124.0.0 :: Published 2019-01-30 :: SQL Workbench/J is a free, DBMS-independent, cross-platform SQL query tool. :: SQL Workbench/J is a free, DBMS-independent, cross-platform SQL query tool. It is written in Java and should run on any operating system that provides a Java Runtime Environment.

"@ | more
}

function Help-ChocoPackaging {
@"
    
cinst -y insted
cinst -y lessmsi          # Analyse and extract MSIs.
cinst -y innosetup
cinst -y nsis
cinst -y wiz35
cinst -y resourcesextract
cinst -y wixtoolset       # Windows Installer XML (WiX) is a toolset that builds Windows installation packages from XML source code. The toolset supports a command line environment that developers may integrate into their build processes to build MSI and MSM setup packages.
"@ | more
}
    
function Help-ChocoMultimedia {
@"

### Video Players
cinst -y mpc-hc-clsid2  # Fork of MPC-HC by clsid2, various new features and updates
cinst -y mpc-hc         # My preferred video player (but deprecated as development stopped in 2018, so I use mpc-hc-clsid2 nwo)
cinst -y mpc-qt         # Reimplementation of MPC-HC in QT
cinst -y vlc potplayer  # Other video players (MPC-HC is better imo)
cinst -y streamlink     # console tool to pipe flash videos to video players like VLC
cinst -y kodi           # Kodi (formerly known as XBMC) media player (Linux, OSX, Windows, iOS and Android)
cinst -y hulu.desktop   # Hulu player
cinst -y plexmediaserver plex-home-theater

### Streaming Apps
cinst -y bbc-iplayer
cinst -y getiplayer
cinst -y opentrack
cinst -y streaming-video-downloader

### Video Editing
cinst -y handbrake ffmpeg virtualdub

### Comic Apps
cinst -y yacreader   # Best Comic Reader on Chocolatey
    # comic-collector comicrack (alternative comic readers, but yacreader is much better imo)

### Image Editing Apps
cinst -y Paint.NET Gimp FireAlPaca PhotoFlare PhotoFlow ImageMagick   # PhotoShop / PaintShopPro are not available in Chocolatey
cinst -y handbrake makemkv mkvtoolnix

### Book Apps
cinst kindlegen -y  # Convert epub to mobi for use on kindles
cinst calibre -y    # Read all E-book epub, mobi, etc
    # https://beebom.com/best-epub-reader-windows/  kobo, nook, cover, bibliovore, none are on chocolatey
    # https://www.microsoft.com/nl-nl/p/freda-epub-ebook-reader/9wzdncrfj43b?ocid=badge&rtc=1&activetab=pivot:overviewtab
    # https://www.amazon.com/kindle-dbs/fd/kcp?tag=georiot-us-default-20&ascsubtag=trd-6368135935388978482-20

### Download/offline YouTube videos
cinst -y youtube-dl
cinst -y win-youtube-dl
cinst -y atubecatcher    
"@ | more
}



# With a given search pattern, run 'choco list', then do 'choco info' against each match.
# Capture PackageName / Version / Summary / Description and format on a single line.
# Collection can take some minutes to complete.
# Usage:
#    choco-get-description.ps1 <searchterm1> <searchterm2> <searchterm3> ...
# or
#    update the items in the $arr_search array and run choco-get-description.ps1 without arguments

# $arr_search = ("docker", "kubernetes", "ansible")
# test if $args is empty, if it is, use $arr_search defined above or populate $arr_search with contents of $args
# if ($numArgs -ne 0) { $arr_search = $args }

####################
#
# ChocoCollect <search1> <search2> <search3> ...
#
####################
function Get-Choco-Descriptions($x) {
    function Confirm-Choice {
        param ( [string]$Message )
        $yes = new-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes";
        $no = new-Object System.Management.Automation.Host.ChoiceDescription "&No", "No";
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no);
        $answer = $host.ui.PromptForChoice($caption, $message, $choices, 0)
        switch ($answer){
            0 {return $true; break}
            1 {return $false; break}
        }
    }
    
    $numArgs = $args.Length
    echo "`nAll args   : $args"
    echo "Total args : $numArgs`n"
    for ($i=0; $i -le $numArgs - 1; $i++) { echo "Argument $i : $($args[$i])" }
    
    $now = (Get-Date).GetDateTimeFormats()[77].ToString()   # [datetime]::Now
    $header = "`nStarted search at $($now) for '$($x)'"
    $filetmp = ".\choco-details__$($x)__list.txt"
    $filetxt_time = Get-Date -format "yyyy-MM-dd" # removed __hh-mm
    $previousfiles = Get-ChildItem ".\choco-details__$($x)__*"
    $previousnum = $previousfiles.count
    
    $exitFunction = $false
    if ($previousnum -ne 0) {
        echo "`r`n`r`nExisting file(s):`r`n$($previousfiles)`r`n`r`n"
        $confirm = "Are you sure you want to delete the previous file(s) matching this search pattern?"
        if (Confirm-Choice $confirm -eq $true) {
            echo "`nrm .\choco-details__$($x)__*   # remove old files"
            rm ".\choco-details__$($x)__*" > $null        # remove old file
            $exitFunction = $true
        }
    }
    if ($exitFunction -eq $true) { return }
    
    echo "$($header)`r`n==============="            # echo "$filetxt `r`n"

    [regex]$exclude = "\[Pending\]|broken|packages|Did you know|Features|https|Chocolatey v|^$|being ignored due to|It is recommended that you reboot|A pending system reboot| validations performed|Validation Warnings"
    # $exclude is to remove reboot warnings, broken packages and junk/empty lines
    #    being ignored due to the current command being used 'list'.
    #    It is recommended that you reboot at your earliest convenience.
    #  - A pending system reboot request has been detected, however, this is
    # 2 validations performed. 1 success(es), 1 warning(s), and 0 error(s).

    (choco list $x | sort) -notmatch $exclude | Out-File $filetmp
    cat $filetmp
    $lines = Get-Content $filetmp
    $num = $lines.count
    if ($num -eq 0) { echo "No matches in chocolatey db.`n"; rm ".\choco-details__$($x)__*" > $null; break }
    $found = $num.ToString() + "_found"
    $filetxt = ".\choco-details__$($x)__$($found)__$($filetxt_time).txt"
    
    $m = "matches"; if ($num -eq 1) {$m = "match" }
    $foundlist = "Found $($num) search $($m) for 'choco list $($x)'"
    
    echo "`r`n`r`n$($mat)`r`n==========`r`n"
    foreach ($line in $lines) {
        $pkg = $line.split(" ")[0]
        $ver = $line.split(" ")[1]
        $pkginfo = choco info $pkg

        # Use try/catch in case Summary or Description are empty, remember ? => Where-Object
        try { $summary = ($pkginfo | ? { $_ -match "Summary:" }).replace(" Summary:"," ::") }
        catch { $summary = " ::" } 
        try { $description = ($pkginfo | ? { $_ -match "Description:" }).replace(" Description:"," ::") }
        catch { $description = " ::" }
        if ($summary -eq $description) { $description = "" }    # ignore description if it is same as summary
        if ($description -match '$summary') { $summary = "" }   # ignore summary if it is contained in the description, note '' around the match otherwise will fail with special characters in $summary e.g. "C++"
        
        # Get Published date and trusted/approved date if available. Note that .split uses chararray, while -split uses string to split
        try {
            $published = (( $pkginfo | ? { $_ -match "Published:" } ) -split " Published: ")[1]
            $pub_day   = ($published -split "/")[0]
            $pub_month = ($published -split "/")[1]
            $pub_year  = ($published -split "/")[2]
            $published = $pub_year + '-' + $pub_month + '-' + $pub_day   # Note that month/day are different way around from $approved so correct to yyyy-mm-dd
        }
        catch { $published = "N.A." }

        try {
            $approved = (( $pkginfo | ? { $_ -match " trusted package on " } ) -split " trusted package on ")[1]
            $approve_month = ($approved -split " ")[0] -replace "Jan", "01" -replace "Feb", "02" -replace "Mar", "03" -replace "Apr", "04" -replace "May", "05" -replace "Jun", "06" -replace "Jul", "07" -replace "Aug", "08" -replace "Sep", "09" -replace "Oct", "10" -replace "Nov", "11" -replace "Dec", "12"
            $approve_day   = ($approved -split " ")[1]
            $approve_year  = ($approved -split " ")[2]
            $approved = $approve_year + '-' + $approve_month + '-' + $approve_day
        }
        catch { $approved = " N.A." } 

        $outpkg = $pkg + " :: " + $ver + " :: Published " + $published + " :: Latest Approval " + $approved + $summary + $description + "`r`n"
        echo $outpkg
        $out = $out + $outpkg + "`r`n"
    }

    $now = (Get-Date).GetDateTimeFormats()[77].ToString()   # [datetime]::Now
    $end = "Ended search at   " + $now + "`r`n"
    echo $header
    echo $end
    echo $foundlist
    echo $header | Out-File $filetxt   # Must not use -Append when creating file or get an error
    echo $end | Out-File -Append $filetxt
    echo $foundlist | Out-File -Append $filetxt
    echo ===============`r`n | Out-File -Append $filetxt    
    $out | Out-File -Append $filetxt
    Remove-Item $filetmp

    foreach ($search in $arr_search) { Get-Choco-Descriptions $search }
}

# Some search terms for chocolatey database:
#    command, commandline, console, sync, replication, backup
#    filesystem, storage, cloud, azure, aws, onenote, dropbox
#    virtualbox, vagrant, puppet, kubernetes, docker, packer, vm, virtual
#    powershell, sql, python, perl, java, boxstarter
#    office, automation, system
#    games, space, stellar, simulator, strategy, arcade, steam, virus



function Choco-Index {
    <#
    .SYNOPSIS
    Choco-Index <search1> <search2> <search3> ...   # Save matches in the chocolatey repository to users Temp folder.
    With <search>, run 'choco list', then 'choco info' against each match, then format all output into single line.
    Captures PackageName / Version / Summary / Description and format all on a single line.
    Results are stored in ps_choco_<search1>.txt in the users Temp folder
    Note that this function has no declared parameters, instead just splitting '$args' to serch on all input strings.
    ToDo: Clean up, and Jobs should greatly speed up processing
    Alternate syntax:   update $arr_search array on console then run cindex.ps1 without arguments
    #>
    if ($null -eq $arr_search) { $arr_seach = @("docker", "kubernetes", "ansible") }

    $numArgs = $args.Length
    ""
    echo "All args   : $args"
    echo "Total args : $numArgs`n"
    for ($i=0; $i -le $numArgs - 1; $i++) { echo "Argument $i : $($args[$i])" }

    # test if $args is empty, if it is, use $arr_search defined above or populate $arr_search with contents of $args
    if ($numArgs -ne 0) { $arr_search = $args }

    function Confirm-Choice {
        param ( [string]$Message )
        $yes = new-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes";
        $no = new-Object System.Management.Automation.Host.ChoiceDescription "&No", "No";
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no);
        $answer = $host.ui.PromptForChoice($caption, $message, $choices, 0)
        switch ($answer){
            0 {return $true; break}
            1 {return $false; break}
        }
    }

    function Get-Choco-Descriptions($x) {
        $now = (Get-Date).GetDateTimeFormats()[77].ToString()   # [datetime]::Now
        $header = "`nStarted search at $($now) for '$($x)'"
        $filetmp = "$env:TEMP\ps_choco-details__$($x)__list.txt"
        $previousfiles = Get-ChildItem "$env:TEMP\choco-details__$($x)__*"
        $previousnum = $previousfiles.count
        $filetxt_time = Get-Date -format "yyyy-MM-dd" # removed __hh-mm

        $exitFunction = $false
        if ($previousnum -ne 0) {
            echo "`r`n`r`nExisting file(s):`r`n$($previousfiles)`r`n`r`n"
            $confirm = "Are you sure you want to delete the previous file(s) matching this search pattern?"
            if (Confirm-Choice $confirm -eq $true) {
                echo "`nrm .\choco-details__$($x)__*   # remove old files"
                rm ".\choco-details__$($x)__*" > $null        # remove old file
                $exitFunction = $true
            }
        }
        if ($exitFunction -eq $true) { return }

        echo "$($header)`r`n==============="            # echo "$filetxt `r`n"

        [regex]$exclude = "\[Pending\]|broken|packages|Did you know|Features|https|Chocolatey v|^$|being ignored due to|It is recommended that you reboot|A pending system reboot| validations performed|Validation Warnings"
        # $exclude is to remove reboot warnings, broken packages and junk/empty lines
        #    being ignored due to the current command being used 'list'.
        #    It is recommended that you reboot at your earliest convenience.
        #  - A pending system reboot request has been detected, however, this is
        # 2 validations performed. 1 success(es), 1 warning(s), and 0 error(s).

        (choco list $x* | sort) -notmatch $exclude | Out-File $filetmp
        cat $filetmp
        $lines = Get-Content $filetmp
        $num = $lines.count
        if ($num -eq 0) { echo "No matches in chocolatey db."; rm ".\choco-details__$($x)__*" > $null; break }
        $found = $num.ToString() + "_found"
        $filetxt = "$env:TEMP\ps_choco-details__$($x)__$($found)__$($filetxt_time).txt"

        $m = "matches"; if ($num -eq 1) {$m = "match" }
        $foundlist = "Found $($num) search $($m) for 'choco list $($x)'"

        echo "`r`n`r`n$($mat)`r`n==========`r`n"
        foreach ($line in $lines) {
            $pkg = $line.split(" ")[0]
            $ver = $line.split(" ")[1]
            $pkginfo = choco info $pkg

            # Use try/catch in case Summary or Description are empty, remember ? => Where-Object
            try { $summary = ($pkginfo | ? { $_ -match "Summary:" }).replace(" Summary:"," ::") }
            catch { $summary = "  ::" } 
            try { $description = ($pkginfo | ? { $_ -match "Description:" }).replace(" Description:"," ::") }
            catch { $description = "  ::" }
            if ($summary -eq $description) { $description = "" }    # ignore description if it is same as summary
            if ($description -match '$summary') { $summary = "" }   # ignore summary if it is contained in the description, note '' around the match otherwise will fail with special characters in $summary e.g. "C++"

            # Get Published date and trusted/approved date if available. Note that .split uses chararray, while -split uses string to split
            try {
                $published = (( $pkginfo | ? { $_ -match "Published:" } ) -split " Published: ")[1]
                $pub_day   = ($published -split "/")[0]
                $pub_month = ($published -split "/")[1]
                $pub_year  = ($published -split "/")[2]
                # month/day are different way around from $approved so correct to yyyy-mm-dd, though this might be locale related
                $published = $pub_year + '-' + $pub_month + '-' + $pub_day
            }
            catch { $published = "--" }
            if ($published -eq "--") { $publishedtext = ""} else { $publishedtext = "$published (P)"}

            try {
                $approved = (( $pkginfo | ? { $_ -match " trusted package on " } ) -split " trusted package on ")[1]
                $approve_month = ($approved -split " ")[0] -replace "Jan", "01" -replace "Feb", "02" -replace "Mar", "03" -replace "Apr", "04" -replace "May", "05" -replace "Jun", "06" -replace "Jul", "07" -replace "Aug", "08" -replace "Sep", "09" -replace "Oct", "10" -replace "Nov", "11" -replace "Dec", "12"
                $approve_day   = ($approved -split " ")[1]
                $approve_year  = ($approved -split " ")[2]
                $approved = $approve_year + '-' + $approve_month + '-' + $approve_day
            }
            catch { $approved = "--" }
            if ($approved -eq "--") { $approvedtext = ""} else { $approvedtext = " $approved (A)"}

            $outpkg = "choco install -y $pkg   # $publishedtext$approvedtext$summary$description`r`n"
            echo $outpkg
            $out = $out + $outpkg + "`r`n"
        }

        $now = (Get-Date).GetDateTimeFormats()[77].ToString()   # [datetime]::Now
        $end = "Ended search at   " + $now + "`r`n"
        echo $header
        echo $end
        echo $foundlist
        echo $header | Out-File $filetxt   # Must not use -Append when creating file or get an error
        echo $end | Out-File -Append $filetxt
        echo $foundlist | Out-File -Append $filetxt
        echo ===============`r`n | Out-File -Append $filetxt    
        $out | Out-File -Append $filetxt
        Remove-Item $filetmp
    }

    foreach ($search in $arr_search) { Get-Choco-Descriptions $search }

    # Some search terms for chocolatey database:
    #    command, commandline, console, sync, replication, backup
    #    filesystem, storage, cloud, azure, aws, onenote, dropbox
    #    virtualbox, vagrant, puppet, kubernetes, docker, packer, vm, virtual
    #    powershell, sql, python, perl, java, boxstarter
    #    office, automation, system
    #    games, space, stellar, simulator, strategy, arcade, steam, virus

    # Note on declaring parameters with preset values and types
    # param( [string]$dir = "C:\Windows", [int32]$size = 200000 )
    #
    # With the above params, the following will test files in arg[0]
    # to see if they are bigger than the value of arg[1]
    #
    # $files = Get-ChildItem $args[0]
    # foreach ($file in $files) { if ($file.length -gt $args[1]) { Write-Output $file }

    # Alternative GetDateTimeFormats (there are 114 standard time format outputs in this)
    # for ($i=1; $i -lt 114; $i++)  { Write-Host "$i :" (Get-Date).GetDateTimeFormats()[$i].ToString() }   # view all TimeFormat
    # (Get-Date).GetDateTimeFormats()[77].ToString()   # 2019-11-19 21:03:20
    #    -replace " ", "__" -replace ":", "-"          # 2019-11-19__21-03-20
    # (Get-Date).GetDateTimeFormats()[57].ToString()   # 2019-11-19 21:03
    #    -replace " ", "__" -replace ":", "-"          # 2019-11-19__21-03
    # (Get-Date).GetDateTimeFormats()[88].ToString()   # 21:03      -replace ":", "-" to make filename compatible
    # (Get-Date).GetDateTimeFormats()[92].ToString()   # 21:03:14   -replace ":", "-" to make filename compatible
}
Set-Alias cindex Choco-Index


function Remove-HiddenAttribute ($Path, [switch]$Recurse) {
    <#
    Purpose: change the Attributes of "Hidden" folder to "Normal"
    website: http://www.amandhally.net/blog
    blog: http://newdelhipowershellusergroup.blogspot.com/               /^(o.o)^\ 
    more info: http://newdelhipowershellusergroup.blogspot.com/2012/01/script-to-reset-hidden-files-and.html 
    #>
    if ($Recurse -eq $true) { $r = "-Recurse" }
    $Items = Get-ChildItem -Path $path -Force  $r
    $HiddenItems = $Filepath | where { $_.Attributes -match "Hidden"}
    $HiddenItems                                                         # Show the items
    foreach ( $Item in $HiddenItems ) { $Item.Attributes = "Archive" }   # Unhide the hidden items
}

# This is old, non-working variant: https://gallery.technet.microsoft.com/scriptcenter/b66434f1-4b3f-4a94-8dc3-e406eb30b750
# Win 10 1903 working version: https://pinto10blog.wordpress.com/2016/09/10/pinto10/
function Masquerade-PEB {
    <#
    .SYNOPSIS
        Masquerade-PEB uses NtQueryInformationProcess to get a handle to powershell's
        PEB. From there it replaces a number of UNICODE_STRING structs in memory to
        give powershell the appearance of a different process. Specifically, the
        function will overwrite powershell's "ImagePathName" & "CommandLine" in
        _RTL_USER_PROCESS_PARAMETERS and the "FullDllName" & "BaseDllName" in the
        _LDR_DATA_TABLE_ENTRY linked list.
        
        This can be useful as it would fool any Windows work-flows which rely solely
        on the Process Status API to check process identity. A practical example would
        be the IFileOperation COM Object which can perform an elevated file copy if it
        thinks powershell is really explorer.exe ;)!
    
        Notes:
          * Works on x32/64.
        
          * Most of these API's and structs are undocumented. I strongly recommend
            @rwfpl's terminus project as a reference guide!
              + http://terminus.rewolf.pl/terminus/
        
          * Masquerade-PEB is basically a reimplementation of two functions in UACME
            by @hFireF0X. My code is quite different because,  unfortunately, I don't
            have access to all those c++ goodies and I could not get a callback for
            LdrEnumerateLoadedModules working!
              + supMasqueradeProcess: https://github.com/hfiref0x/UACME/blob/master/Source/Akagi/sup.c#L504
              + supxLdrEnumModulesCallback: https://github.com/hfiref0x/UACME/blob/master/Source/Akagi/sup.c#L477
    
    .DESCRIPTION
        Author: Ruben Boonen (@FuzzySec)
        License: BSD 3-Clause
        Required Dependencies: None
        Optional Dependencies: None
    
    .EXAMPLE
        C:\PS> Masquerade-PEB -BinPath "C:\Windows\explorer.exe"
        # Run the script with two arguments.  The first is the full path to the file you wish to pin and the second is either PIN or UNPIN.
        # PinToTaskBar1903.ps1 C:\Windows\notepad.exe PIN
    #>
    
    param (
        [Parameter(Mandatory = $True)]
        [string]$BinPath
    )

    Add-Type -TypeDefinition @"
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Security.Principal;

    [StructLayout(LayoutKind.Sequential)]
    public struct UNICODE_STRING
    {
        public UInt16 Length;
        public UInt16 MaximumLength;
        public IntPtr Buffer;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _LIST_ENTRY
    {
        public IntPtr Flink;
        public IntPtr Blink;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct _PROCESS_BASIC_INFORMATION
    {
        public IntPtr ExitStatus;
        public IntPtr PebBaseAddress;
        public IntPtr AffinityMask;
        public IntPtr BasePriority;
        public UIntPtr UniqueProcessId;
        public IntPtr InheritedFromUniqueProcessId;
    }

    /// Partial _PEB
    [StructLayout(LayoutKind.Explicit, Size = 64)]
    public struct _PEB
    {
        [FieldOffset(12)]
        public IntPtr Ldr32;
        [FieldOffset(16)]
        public IntPtr ProcessParameters32;
        [FieldOffset(24)]
        public IntPtr Ldr64;
        [FieldOffset(28)]
        public IntPtr FastPebLock32;
        [FieldOffset(32)]
        public IntPtr ProcessParameters64;
        [FieldOffset(56)]
        public IntPtr FastPebLock64;
    }

    /// Partial _PEB_LDR_DATA
    [StructLayout(LayoutKind.Sequential)]
    public struct _PEB_LDR_DATA
    {
        public UInt32 Length;
        public Byte Initialized;
        public IntPtr SsHandle;
        public _LIST_ENTRY InLoadOrderModuleList;
        public _LIST_ENTRY InMemoryOrderModuleList;
        public _LIST_ENTRY InInitializationOrderModuleList;
        public IntPtr EntryInProgress;
    }

    /// Partial _LDR_DATA_TABLE_ENTRY
    [StructLayout(LayoutKind.Sequential)]
    public struct _LDR_DATA_TABLE_ENTRY
    {
        public _LIST_ENTRY InLoadOrderLinks;
        public _LIST_ENTRY InMemoryOrderLinks;
        public _LIST_ENTRY InInitializationOrderLinks;
        public IntPtr DllBase;
        public IntPtr EntryPoint;
        public UInt32 SizeOfImage;
        public UNICODE_STRING FullDllName;
        public UNICODE_STRING BaseDllName;
    }

    public static class Kernel32
    {
        [DllImport("kernel32.dll")]
        public static extern UInt32 GetLastError();

        [DllImport("kernel32.dll")]
        public static extern Boolean VirtualProtectEx(
            IntPtr hProcess,
            IntPtr lpAddress,
            UInt32 dwSize,
            UInt32 flNewProtect,
            ref UInt32 lpflOldProtect);

        [DllImport("kernel32.dll")]
        public static extern Boolean WriteProcessMemory(
            IntPtr hProcess,
            IntPtr lpBaseAddress,
            IntPtr lpBuffer,
            UInt32 nSize,
            ref UInt32 lpNumberOfBytesWritten);
    }

    public static class Ntdll
    {
        [DllImport("ntdll.dll")]
        public static extern int NtQueryInformationProcess(
            IntPtr processHandle, 
            int processInformationClass,
            ref _PROCESS_BASIC_INFORMATION processInformation,
            int processInformationLength,
            ref int returnLength);

        [DllImport("ntdll.dll")]
        public static extern void RtlEnterCriticalSection(
            IntPtr lpCriticalSection);

        [DllImport("ntdll.dll")]
        public static extern void RtlLeaveCriticalSection(
            IntPtr lpCriticalSection);
    }
"@
    
    # Flag architecture $x32Architecture/!$x32Architecture
    if ([System.IntPtr]::Size -eq 4) {
        $x32Architecture = 1
    }

    # Current Proc handle
    $ProcHandle = (Get-Process -Id ([System.Diagnostics.Process]::GetCurrentProcess().Id)).Handle

    # Helper function to overwrite UNICODE_STRING structs in memory
    function Emit-UNICODE_STRING {
        param(
            [IntPtr]$hProcess,
            [IntPtr]$lpBaseAddress,
            [UInt32]$dwSize,
            [String]$data
        )

        # Set access protections -> PAGE_EXECUTE_READWRITE
        [UInt32]$lpflOldProtect = 0
        $CallResult = [Kernel32]::VirtualProtectEx($hProcess, $lpBaseAddress, $dwSize, 0x40, [ref]$lpflOldProtect)

        # Create replacement struct
        $UnicodeObject = New-Object UNICODE_STRING
        $UnicodeObject_Buffer = $data
        [UInt16]$UnicodeObject.Length = $UnicodeObject_Buffer.Length*2
        [UInt16]$UnicodeObject.MaximumLength = $UnicodeObject.Length+1
        [IntPtr]$UnicodeObject.Buffer = [System.Runtime.InteropServices.Marshal]::StringToHGlobalUni($UnicodeObject_Buffer)
        [IntPtr]$InMemoryStruct = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($dwSize)
        [system.runtime.interopservices.marshal]::StructureToPtr($UnicodeObject, $InMemoryStruct, $true)

        # Overwrite PEB UNICODE_STRING struct
        [UInt32]$lpNumberOfBytesWritten = 0
        $CallResult = [Kernel32]::WriteProcessMemory($hProcess, $lpBaseAddress, $InMemoryStruct, $dwSize, [ref]$lpNumberOfBytesWritten)

        # Free $InMemoryStruct
        [System.Runtime.InteropServices.Marshal]::FreeHGlobal($InMemoryStruct)
    }

    # Process Basic Information
    $PROCESS_BASIC_INFORMATION = New-Object _PROCESS_BASIC_INFORMATION
    $PROCESS_BASIC_INFORMATION_Size = [System.Runtime.InteropServices.Marshal]::SizeOf($PROCESS_BASIC_INFORMATION)
    $returnLength = New-Object Int
    $CallResult = [Ntdll]::NtQueryInformationProcess($ProcHandle, 0, [ref]$PROCESS_BASIC_INFORMATION, $PROCESS_BASIC_INFORMATION_Size, [ref]$returnLength)

    # PID & PEB address
    # echo "`n[?] PID $($PROCESS_BASIC_INFORMATION.UniqueProcessId)"
    if ($x32Architecture) {
        # echo "[+] PebBaseAddress: 0x$("{0:X8}" -f $PROCESS_BASIC_INFORMATION.PebBaseAddress.ToInt32())"
    } else {
        # echo "[+] PebBaseAddress: 0x$("{0:X16}" -f $PROCESS_BASIC_INFORMATION.PebBaseAddress.ToInt64())"
    }

    # Lazy PEB parsing
    $_PEB = New-Object _PEB
    $_PEB = $_PEB.GetType()
    $BufferOffset = $PROCESS_BASIC_INFORMATION.PebBaseAddress.ToInt64()
    $NewIntPtr = New-Object System.Intptr -ArgumentList $BufferOffset
    $PEBFlags = [system.runtime.interopservices.marshal]::PtrToStructure($NewIntPtr, [type]$_PEB)

    # Take ownership of PEB
    # Not sure this is strictly necessary but why not!
    if ($x32Architecture) {
        [Ntdll]::RtlEnterCriticalSection($PEBFlags.FastPebLock32)
    } else {
        [Ntdll]::RtlEnterCriticalSection($PEBFlags.FastPebLock64)
    } # echo "[!] RtlEnterCriticalSection --> &Peb->FastPebLock"

    # &Peb->ProcessParameters->ImagePathName/CommandLine
    if ($x32Architecture) {
        # Offset to &Peb->ProcessParameters
        $PROCESS_PARAMETERS = $PEBFlags.ProcessParameters32.ToInt64()
        # x86 UNICODE_STRING struct's --> Size 8-bytes = (UInt16*2)+IntPtr
        [UInt32]$StructSize = 8
        $ImagePathName = $PROCESS_PARAMETERS + 0x38
        $CommandLine = $PROCESS_PARAMETERS + 0x40
    } else {
        # Offset to &Peb->ProcessParameters
        $PROCESS_PARAMETERS = $PEBFlags.ProcessParameters64.ToInt64()
        # x64 UNICODE_STRING struct's --> Size 16-bytes = (UInt16*2)+IntPtr
        [UInt32]$StructSize = 16
        $ImagePathName = $PROCESS_PARAMETERS + 0x60
        $CommandLine = $PROCESS_PARAMETERS + 0x70
    }

    # Overwrite PEB struct
    # Can easily be extended to other UNICODE_STRING structs in _RTL_USER_PROCESS_PARAMETERS(/or in general)
    $ImagePathNamePtr = New-Object System.Intptr -ArgumentList $ImagePathName
    $CommandLinePtr = New-Object System.Intptr -ArgumentList $CommandLine
    if ($x32Architecture) {
        # echo "[>] Overwriting &Peb->ProcessParameters.ImagePathName: 0x$("{0:X8}" -f $ImagePathName)"
        # echo "[>] Overwriting &Peb->ProcessParameters.CommandLine: 0x$("{0:X8}" -f $CommandLine)"
    } else {
        # echo "[>] Overwriting &Peb->ProcessParameters.ImagePathName: 0x$("{0:X16}" -f $ImagePathName)"
        # echo "[>] Overwriting &Peb->ProcessParameters.CommandLine: 0x$("{0:X16}" -f $CommandLine)"
    }
    Emit-UNICODE_STRING -hProcess $ProcHandle -lpBaseAddress $ImagePathNamePtr -dwSize $StructSize -data $BinPath
    Emit-UNICODE_STRING -hProcess $ProcHandle -lpBaseAddress $CommandLinePtr -dwSize $StructSize -data $BinPath

    # &Peb->Ldr
    $_PEB_LDR_DATA = New-Object _PEB_LDR_DATA
    $_PEB_LDR_DATA = $_PEB_LDR_DATA.GetType()
    if ($x32Architecture) {
        $BufferOffset = $PEBFlags.Ldr32.ToInt64()
    } else {
        $BufferOffset = $PEBFlags.Ldr64.ToInt64()
    }
    $NewIntPtr = New-Object System.Intptr -ArgumentList $BufferOffset
    $LDRFlags = [system.runtime.interopservices.marshal]::PtrToStructure($NewIntPtr, [type]$_PEB_LDR_DATA)

    # &Peb->Ldr->InLoadOrderModuleList->Flink
    $_LDR_DATA_TABLE_ENTRY = New-Object _LDR_DATA_TABLE_ENTRY
    $_LDR_DATA_TABLE_ENTRY = $_LDR_DATA_TABLE_ENTRY.GetType()
    $BufferOffset = $LDRFlags.InLoadOrderModuleList.Flink.ToInt64()
    $NewIntPtr = New-Object System.Intptr -ArgumentList $BufferOffset

    # Traverse doubly linked list
    # &Peb->Ldr->InLoadOrderModuleList->InLoadOrderLinks->Flink
    # This is probably overkill, powershell.exe should always be the first entry for InLoadOrderLinks
    # echo "[?] Traversing &Peb->Ldr->InLoadOrderModuleList doubly linked list"
    while ($ListIndex -ne $LDRFlags.InLoadOrderModuleList.Blink) {
        $LDREntry = [system.runtime.interopservices.marshal]::PtrToStructure($NewIntPtr, [type]$_LDR_DATA_TABLE_ENTRY)

        if ([System.Runtime.InteropServices.Marshal]::PtrToStringUni($LDREntry.FullDllName.Buffer) -like "*powershell.exe*") {

            if ($x32Architecture) {
                # x86 UNICODE_STRING struct's --> Size 8-bytes = (UInt16*2)+IntPtr
                [UInt32]$StructSize = 8
                $FullDllName = $BufferOffset + 0x24
                $BaseDllName = $BufferOffset + 0x2C
            } else {
                # x64 UNICODE_STRING struct's --> Size 16-bytes = (UInt16*2)+IntPtr
                [UInt32]$StructSize = 16
                $FullDllName = $BufferOffset + 0x48
                $BaseDllName = $BufferOffset + 0x58
            }

            # Overwrite _LDR_DATA_TABLE_ENTRY struct
            # Can easily be extended to other UNICODE_STRING structs in _LDR_DATA_TABLE_ENTRY(/or in general)
            $FullDllNamePtr = New-Object System.Intptr -ArgumentList $FullDllName
            $BaseDllNamePtr = New-Object System.Intptr -ArgumentList $BaseDllName
            if ($x32Architecture) {
                # echo "[>] Overwriting _LDR_DATA_TABLE_ENTRY.FullDllName: 0x$("{0:X8}" -f $FullDllName)"
                # echo "[>] Overwriting _LDR_DATA_TABLE_ENTRY.BaseDllName: 0x$("{0:X8}" -f $BaseDllName)"
            } else {
                # echo "[>] Overwriting _LDR_DATA_TABLE_ENTRY.FullDllName: 0x$("{0:X16}" -f $FullDllName)"
                # echo "[>] Overwriting _LDR_DATA_TABLE_ENTRY.BaseDllName: 0x$("{0:X16}" -f $BaseDllName)"
            }
            Emit-UNICODE_STRING -hProcess $ProcHandle -lpBaseAddress $FullDllNamePtr -dwSize $StructSize -data $BinPath
            Emit-UNICODE_STRING -hProcess $ProcHandle -lpBaseAddress $BaseDllNamePtr -dwSize $StructSize -data $BinPath
        }
        
        $ListIndex = $BufferOffset = $LDREntry.InLoadOrderLinks.Flink.ToInt64()
        $NewIntPtr = New-Object System.Intptr -ArgumentList $BufferOffset
    }

    # Release ownership of PEB
    if ($x32Architecture) {
        [Ntdll]::RtlLeaveCriticalSection($PEBFlags.FastPebLock32)
    } else {
        [Ntdll]::RtlLeaveCriticalSection($PEBFlags.FastPebLock64)
    } # echo "[!] RtlLeaveCriticalSection --> &Peb->FastPebLock`n"
}
    
function PinToTaskBar {    
    if (($args[0] -eq "/?") -Or ($args[0] -eq "-h") -Or ($args[0] -eq "--h") -Or ($args[0] -eq "-help") -Or ($args[0] -eq "--help")){
        write-host "This script needs to be run with two arguments."`r`n
        write-host "1 - Full path to the file you wish to pin (surround in quotes)."
        write-host "2 - Either PIN or UNPIN (case insensitive)."
        write-host "Example:-"
        # write-host 'powershell -noprofile -ExecutionPolicy Bypass -file PinToTaskBar1903.ps1 "C:\Windows\Notepad.exe" PIN'`r`n
        write-host 'PinToTaskBar "C:\Windows\Notepad.exe" PIN'`r`n
        Break
    }

    if ($args.count -eq 2){
        $TargetFile = $args[0]
        $PinUnpin = $args[1].ToUpper()
    } else {
        write-host "Incorrect number of arguments.  Exiting..."`r`n
        Break
    }

    # Check all the variables are correct before starting
    if (!(Test-Path "$TargetFile")){
        write-host "File not found.  Exiting..."`r`n
        Break
    }

    if (($PinUnpin -ne "PIN") -And ($PinUnpin -ne "UNPIN")){
        write-host "Second argument not set to PIN or UNPIN.  Exiting..."`r`n
        Break
    }

    # Set the arguments to the required verb actions
    if ($PinUnpin -eq "PIN"){$PinUnpin = "taskbarpin"}
    if ($PinUnpin -eq "UNPIN"){$PinUnpin = "taskbarunpin"}

    # Split the target path to folder, filename and filename with no extension
    $FileNameNoExt = (Get-ChildItem $TargetFile).BaseName
    $FileNameWithExt = (Get-ChildItem $TargetFile).Name
    $Directory = (Get-Childitem $TargetFile).Directory

    # Hide Powershell as Explorer...
    Masquerade-PEB -BinPath "C:\Windows\explorer.exe"

    # If target file is not a .lnk then create a shortcut, (un)pin that and then delete it
    if ((Get-ChildItem $TargetFile).Extension -ne ".lnk"){

        if (test-path "$env:TEMP\$FileNameNoExt.lnk"){Remove-Item -path "$env:TEMP\$FileNameNoExt.lnk"}

        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$env:TEMP\$FileNameNoExt.lnk")
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()

        $TargetFile = "$env:TEMP\$FileNameNoExt.lnk"
        $FileNameWithExt = (Get-ChildItem $TargetFile).Name
        $Directory = (Get-Childitem $TargetFile).Directory

        (New-Object -ComObject shell.application).Namespace("$Directory\").parsename("$FileNameWithExt").invokeverb("$PinUnpin")

        if (test-path "$env:TEMP\$FileNameNoExt.lnk"){Remove-Item -path "$env:TEMP\$FileNameNoExt.lnk"}

    } else {
        (New-Object -ComObject shell.application).Namespace("$Directory\").parsename("$FileNameWithExt").invokeverb("$PinUnpin")
    }   
}

function Download-Script ($url, $FinalName) {
    if ($url -eq '') { "Require URL to perform download." ; break}
    $DownloadName = ($url -split "/")[-1]   # Could also use:  $url -split "/" | select -last 1   # 'hi there, how are you' -split '\s+' | select -last 1
    $OutPath = Join-Path $ScriptsPath $DownloadName
    "Checking for '$OutPath' ..."
    if ($null -eq $FinalName) { $FinalName = $DownloadName }
    if (!(Test-Path "$ScriptsPath\$DownloadName") -and !(Test-Path "$ScriptName\$FinalName")) {
        "Downloading '$DownloadName' to '$OutPath' ..."
        try { (New-Object System.Net.WebClient).DownloadString($url) | Out-File $OutPath }
        catch { "Failed to download $FileName. Check internet connection, particularly TLS / VPN." }
        if ($FinalName -ne "") {
            if (Test-Path "$ScriptsPath\$DownloadName") {
                if (Test-Path "$ScriptPath\$FinalName") { Move-Item }
                Move-Item "$ScriptsPath\$DownloadName" "$ScriptsPath\$FinalName" -Force
                "Renamed '$DownloadName' to '$FinalName'" 
           }
        }
    }
}

function Download-ScriptExamples {
    # Download various scripts to the $Profile Scripts folder
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host '4. Downoad latest versions of online scripts.' -ForegroundColor Green
    Write-Host ''
    Write-Host '   Place in the default PowerShell script folder and add to path' -ForegroundColor Yellow
    Write-Host '   making them fully usable in any console. Can add more scripts' -ForegroundColor Yellow
    Write-Host '   easily to be deployed to all systems at setup as required.' -ForegroundColor Yellow
    Write-Host "   C:\Users\$HomeLeaf\Documents\WindowsPowerShell\Scripts" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Green
    Write-Host ""

    # Can also install scripts with the *-Script Cmdlets
    # Install-Script, Find-Script, Publish-Script, Save-Script, Uninstall-Script, Update-Script
    # -Scope AllUsers    => C:\Program Files\WindowsPowerShell\Scripts
    # -Scope CurrentUser => $HomeFix\Documents\WindowsPowerShell\Scripts *or* ING VPN
    # Due to the Corporate VPN issue, use the "Find-Script | Save-Script trick"
    $UserScripts = "C:\Users\$HomeLeaf\Documents\WindowsPowerShell\Scripts"
    if (!(Test-Path $UserScripts)) { New-Item $UserScripts -ItemType Directory -Force }


    # Update function ($scriptname, $url) where scriptname will be the final name of the file on disk
    # Update to not perform the download if the script is already there! Notify if so
    # ToDo: Archive all downloaded scripts in case links are broken
    # ToDo: add try / throw for failure on download

    Download-Script 'https://gallery.technet.microsoft.com/scriptcenter/Powershell-function-to-add-a7ac5229/file/166758/1/Add-Path.ps1'
    # https://superwidgets.wordpress.com/2017/01/04/powershell-script-to-report-on-computer-inventory/

    Download-Script 'https://gallery.technet.microsoft.com/scriptcenter/Powershell-Script-to-ping-15e0610a/file/127965/4/Ping-Report-v3.ps1' 'Ping-Report-v3.ps1' 
    # if (Test-Path "$ScriptsPath\Ping-Report-v3.ps1") { Move-Item "$ScriptsPath\Ping-Report-v3.ps1" "$ScriptsPath\Ping-Report.ps1" -Force }   # remove the -v3 from filename

    Download-Script 'https://gallery.technet.microsoft.com/scriptcenter/Fast-asynchronous-ping-IP-d0a5cf0e/file/124575/1/Ping-IPrange.ps1'

    # mklement0 Tools: https://gist.github.com/mklement0/146f3202a810a74cb54a2d353ee4003f
    #    function Show-OperatorHelp { / function Show-TypeHelp { , Shows documentation for built-in .NET types, etc
    Download-Script 'https://gist.githubusercontent.com/mklement0/146f3202a810a74cb54a2d353ee4003f/raw/044746494a61c212cad196a1a12c086e826ba719/Show-OperatorHelp.ps1'
    Download-Script 'https://gist.githubusercontent.com/mklement0/50a1b101cd53978cd147b4b138fe6ef4/raw/9c4dfd2878dfdf8d74eccae707183cdfe536f436/Show-TypeHelp.ps1'

    # Set-Window.ps1
    Download-Script 'https://gallery.technet.microsoft.com/scriptcenter/Set-the-position-and-size-54853527/file/146291/1/Set-Window.ps1'

    # Get-RegistryKeyLastWriteTime.ps1
    Download-Script 'https://gallery.technet.microsoft.com/scriptcenter/Get-RegistryKeyLastWriteTim-63f4dd96/file/131343/1/Get-RegistryKeyLastWriteTime.ps1'
    # Get-RegistryKey.ps1 using 'Find-Script' Cmdlet
    if (!(Test-Path "$ScriptsPath\Get-RegistryKey.ps1")) { try { Find-Script Get-RegistryKey | Save-Script -Path $ScriptName } catch { "Could not connect to PSGallery for Get-ReistryKey.ps1"} }

    # Invoke-TaskCleanerBypass to run an elevated task and skip the UAC prompt (by using task scheduler)
    Download-Script 'https://raw.githubusercontent.com/PoSHMagiC0de/Invoke-TaskCleanerBypass/master/Invoke-TaskCleanerBypass.ps1'

    # Console Art Example
    # https://powershell.org/forums/topic/variables-in-write-host-command/
    Download-Script 'https://gist.github.com/shanenin/f164c483db513b88ce91/raw' 'ConsoleArt.ps1'

    Download-Script 'https://gallery.technet.microsoft.com/scriptcenter/Collects-Remote-Computer-40e2a300/file/171821/1/DesktopInventory.PS1' ''
    Download-Script 'https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Hardware-f99336f6/file/188552/2/Get-Inventory.ps1' ''

    "`nTo update the above scripts, remove old versions from '$ScriptsPath' then rerun this script.`n"

    ### First run of Install-Script does the following:
    # PATH Environment Variable Change
    # Your system has not been configured with a default script installation path yet, which means you can only run a
    # script by specifying the full path to the script file. This action places the script into the folder 'C:\Program
    # Files\WindowsPowerShell\Scripts', and adds that folder to your PATH environment variable. Do you want to add the
    # script installation path 'C:\Program Files\WindowsPowerShell\Scripts' to the PATH environment variable?
    # [Y] Yes  [N] No  [S] Suspend  [?] Help (default is "Y"): y
}

function Install-ModuleToDirectory {
    [CmdletBinding()] [OutputType('System.Management.Automation.PSModuleInfo')]
    param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]                                    $Name,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidateScript({ Test-Path $_ })] $Destination
    )

    # Force installation to User Modules, even if running as Admin. i.e. keep all installations in User space.
    # Default user space location for Modules is here, but if Admin, it will try to install to C:\Program Files\WindowsPowerShell
    # We want to force installation into user space, but there is no command to do this
    # There is a command to force installation to AllUser space: Install-Module <Name> -Scope AllUsers
    # $ModulePath = "$(Split-Path $Profile)\Modules1"   # This is where all modules must go. It should be on path (must add if required)
    # $ModulePathTest = foreach ($i in ($env:PSModulePath).split(";")) { if ($i -like $ModulePath) { $True } }   # Get 
    # if ($ModulePathTest -eq $null) { }   # Need to add this if not present

    if (!(Test-Path $UserModulePath)) { New-Item $UserModulePath -ItemType Directory -Force }

    if (($Profile -like "\\*") -and (Test-Path (Join-Path $UserModulePath $Name))) {
        if (Test-Administrator -eq $true) {
            "remove module from network share and move to $Destination"
            # Nothing in here will happen unless working on laptop with a network share
            Uninstall-Module $Name -Force -Verbose
            # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share if in use so use Save-Module
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
        }
        else {
            Write-Host "Module found on network share module path but need to be administrator and connected to VPN" -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "to correctly move Modules into the users module folder on C:\" -ForegroundColor Yellow -BackgroundColor Black
        }
    }
    elseif (Test-Path (Join-Path $AdminModulesPath $Name)) {
        if (Test-Administrator -eq $true) {
            "remove module from $AdminModulesPath and move to $Destination"
            Uninstall-Module $Name -Force -Verbose
            # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share if in use so use Save-Module
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
        }
        else {
            Write-Host "Module found on in Admin Modules folder: $(split-path $AdminModulesPath) C:\Program Files\WindowsPowerShell\Modules." -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "Need to be Admin to correctly move Modules into the users module folder on C:\" -ForegroundColor Yellow -BackgroundColor Black
        }
    }
    # Get-InstalledModule   # Shows only the Modules installed by PowerShellGet.
    # Get-Module            # Gets the modules that have been imported or that can be imported into the current session.
    elseif (Test-Path (Join-Path $Destination $Name)) {
        # https://stackoverflow.com/questions/48424152/compare-system-version-in-powershell
        # To use the repository, you either need PowerShell 5 or install the PowerShellGet module manually (which is
        # available for download on powershellgallery.com) to get Find/Save/Install/Update/Remove-Script for Modules.
        # https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/getting-latest-powershell-gallery-module-version
        # https://stackoverflow.com/questions/52633919/powershell-sort-version-objects-descending
        # "1.." -match "\b\d(\.\d{0,5}){0,3}\d$"
        # https://techibee.com/powershell/check-if-a-string-contains-numbers-in-it-using-powershell/2842
        $ModVerLocal = (Get-Module $Name -ListAvailable -EA Silent).Version
        $ModVerOnline = Get-PublishedModuleVersion $Name
        $ModVerLocal = "$(($ModVerLocal).Major).$(($ModVerLocal).Minor).$(($ModVerLocal).Build)"      # reuse the [version] variable as a [string]
        $ModVerOnline = "$(($ModverOnline).Major).$(($ModverOnline).Minor).$(($ModverOnline).Build)"  # reuse the [version] variable as a [string]
        # if ($ModuleVersionOnline -ne "") { $ModuleVersionOnline = "$($ModuleVersionOnline.split(".")[0]).$($ModuleVersionOnline.split(".")[1]).$($ModuleVersionOnline.split(".")[2])" }
        echo "Local Version:  $ModVerLocal"
        echo "Online Version: $ModVerOnline"
        if ($ModVerLocal -eq $ModVerOnline) {
            echo "$Name is installed and latest version, nothing to do!"
        }
        else {
            if ([bool](Get-Module $Name) -eq $true) { Uninstall-Module $Name -Force -Verbose }
            rm (Join-Path $Destination $Name) -Force -Recurse -Verbose
            Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination -Force -Verbose   # Install the module to the custom destination.
            Import-Module -FullyQualifiedName (Join-Path $Destination $Name) -Force -Verbose
        }
    }
    else {   # Final case is no module is in network share, or local admin modules, or local user modules so now just install it
        Get-PublishedModuleVersion $Name
        Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination -Force -Verbose   # Install the module to the custom destination.
        Import-Module -FullyQualifiedName (Join-Path $Destination $Name) -Force -Verbose
    }

    # Finally, output the Path to the newly installed module and the functions contained in it
    (Get-Module $Name | select Path).Path
    $out = ""; foreach ($i in (Get-Command -Module $Name).Name) {$out = "$out, $i"} ; "" ; Write-Wrap $out.trimstart(", ") ; ""
    # return (Get-Module)
}

# function Install-ModuleToDirectory {
#     [CmdletBinding()] [OutputType('System.Management.Automation.PSModuleInfo')]
#     param(
#         [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]                                    $Name,
#         [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidateScript({ Test-Path $_ })] $Destination
#     )
# 
#     if (($Profile -like "\\*") -and (Test-Path (Join-Path $UserModulePath $Name))) {
#         if (Test-Administrator -eq $true) {
#             # Nothing in here will happen unless working on laptop with a network share
#             Uninstall-Module $Name -Force -Verbose
#             # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!) so use Save-Module
#             Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
#             Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
#         }
#         else {
#             "Module found on network share module path but need to be administrator and connected to VPN"
#             "to correctly move Modules into the users module folder on C:\"
#             pause
#         }
#     }
#     elseif (($Profile -like "\\*") -and (Test-Path (Join-Path $Profile $Name))) {
#         if (Test-Administrator -eq $true) {
#             # Nothing in here will happen unless working on laptop with a network share
#             Uninstall-Module $Name -Force -Verbose
#             # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!) so use Save-Module
#             Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
#             Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
#         }
#         else {
#             "Module found on network share module path but need to be administrator and connected to VPN"
#             "to correctly move Modules into the users module folder on C:\"
#             pause
#         }
#     }
#     elseif (Test-Path (Join-Path $AdminModulePath $Name)) {
#         if (Test-Administrator -eq $true) {
#             # Nothing in here will happen unless working on laptop with a network share
#             Uninstall-Module $Name -Force -Verbose
#             # Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!) so use Save-Module
#             Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
#             Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
#         }
#         else {
#             "Module found on in Admin Modules folder: C:\Program Files\WindowsPowerShell\Modules."
#             "Need to be Admin to correctly move Modules into the users module folder on C:\"
#             pause
#         }
#     }
#     else {
#         Find-Module -Name $Name -Repository 'PSGallery' | Save-Module -Path $Destination   # Install the module to the custom destination.
#         Import-Module -FullyQualifiedName (Join-Path $Destination $Name)
#     }
#     $out = ""; foreach ($i in (Get-Command -Module $Name).Name) {$out = "$out, $i"} ; "" ; Write-Wrap $x.trimstart(", ") ; ""
#     # return (Get-Module)
# }
# 
# Install-ModuleToDirectory -Name 'XXX' -Destination 'E:\Modules'
# try {
#     # Note additional switches if required: -Repository $MyRepoName -Credential $Credential
#     # If the module is already installed, use Update, otherwise use Install
#     if ([bool](Get-Module $Name -ListAvailable)) {
#          Update-Module $Name -Verbose -ErrorAction Stop 
#     } else {
#          Install-Module $Name -Scope CurrentUser -Verbose -ErrorAction Stop
#     }
# } catch {
#     # But if something went wrong, just -Force it, hard.
#     Install-Module $Name -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
# }
#
# $ModuleNetShare = "$(Split-Path $ProfileNetShare)\Modules\$MyModule"
# $ModuleCProfile = "C:\Users\$env:Username\Documents\WindowsPowerShell\Modules\$MyModule"
# $ModuleNetShare
# $ModuleCProfile
#
#     $success = 0
#     # if ($null -ne $(Test-Path $ModuleNetShare)) {
#     # Only run this if $Profile is pointing at Net Share
#     # Note that the uninstalls will fail unless connected to the VPN!
#     if ((Test-Path $ModuleNetShare) -and ($Profile -like "\\*")) {
#         if (Test-Administrator -eq $true) {
#             # Nothing in here will happen unless working on laptop with a network share
#             # First uninstall, then reinstall to get latest version, then move it to $Profile
#             Uninstall-Module $MyModule -Force -Verbose
#             Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!)
#             Move-Item $ModuleNetShare $ModuleCProfile -Force -Verbose
#             Uninstall-Module $MyModule -Force -Verbose                    # Need to uninstall again to clear network share reference from registry
#             Import-Module $MyModule -Scope Local -Force -Verbose          # Finally, import the version in C:\
#             $success = 1
#         }
#         else {
#             "Module found on network share module path but need to be administrator and connected to VPN"
#             "to correctly move Modules into the users module folder on C:\"
#             pause
#         }
#     }
#     if ((Test-Path "C:\Program Files\WindowsPowerShell\Modules\$MyModule") -and (Test-Administrator -eq $true)) {
#         # This is if the module has been loaded into the Administrator folder.
#         # This will move it to the user folder and update
#         Uninstall-Module $MyModule -Force -Verbose
#         Install-Module $MyModule -Scope CurrentUser -Force -Verbose   # This will *always* install to network share(!), so same situation as before
#         Move-Item $ModuleNetShare $ModuleCProfile -Force -Verbose
#         Uninstall-Module $MyModule -Force -Verbose                    # Need to uninstall again to clear network share reference from registry
#         Import-Module $MyModule -Scope Local -Force -Verbose          # Finally, import the version in C:\
#         $success = 1
#     }
#     else {
#         "Module $MyModule found in 'C:\Program Files\WindowsPowerShell\Modules' but need to be administrator"
#         "to correctly move Modules into the users module folder on C:\"
#         pause
#     }
# 
#     if ($success -eq 0) {
#         try {
#             # Note additional switches if required: -Repository $MyRepoName -Credential $Credential
#             # If the module is already installed, use Update, otherwise use Install
#             if ([bool](Get-Module $MyModule -ListAvailable)) {
#                  Update-Module $MyModule -Verbose -ErrorAction Stop 
#             } else {
#                  Install-Module $MyModule -Scope CurrentUser -Verbose -ErrorAction Stop
#             }
#         } catch {
#             # But if something went wrong, just -Force it, hard.
#             Install-Module $MyModule -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
#         }
#     }
# }

function Install-ModuleExamples {
    Write-Host ''
    Write-Host ''
    Write-Host "Setup some useful PowerShell Gallery Modules" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "============================================" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host ''
    Write-Host ''

    # $dependencies = @{   # Taken this hash table and Module installer from https://github.com/pauby/PSTodoWarrior/blob/master/build.ps1
    #     InvokeBuild         = 'latest'
    #     Configuration       = 'latest'
    #     PowerShellBuild     = 'latest'
    #     Pester              = 'latest'
    #     PSScriptAnalyzer    = 'latest'
    #     PSPesterTestHelpers = 'latest'
    #     PSDeploy            = 'latest'  # Maybe pin the version in case he breaks this...
    #     PSTodoTxt           = 'latest'
    # }
    # 
    # # Dependencies
    # if (-not (Get-Command -Name 'Get-PackageProvider' -EA silent)) {
    #     $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    #     Write-Verbose 'Bootstrapping NuGet package provider.'
    #     Get-PackageProvider -Name NuGet -ForceBootstrap | Out-Null
    # } elseif ((Get-PackageProvider).Name -notcontains 'nuget') {
    #     Write-Verbose 'Bootstrapping NuGet package provider.'
    #     Get-PackageProvider -Name NuGet -ForceBootstrap
    # }
    # 
    # # Trust the PSGallery is needed
    # if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
    #     Write-Verbose "Trusting PowerShellGallery."
    #     Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    # }

    Write-Host ''
    Write-Host "Install-Module PSReadLine" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "   Out-Default"
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module psreadline" -ForegroundColor Yellow
    Write-Host "Note: this is installed by default on Windows 10 but not on Windows 7. It is required for many"
    Write-Host "console functions:"
    Write-Host " - History searches with Ctrl+R."
    Write-Host " - Type part of a command then F8 to see go to last matching command."
    Write-Host " - Ctrl+Alt+Shif+? to see all PSReadLine shortcuts."
    Write-Host "For Windows 10, there is nothing to do, but for Windows 7 (even with PowerShell 5.1)"
    Write-Host "it must also be loaded into every session (unlike Windows 10 where it loads by default)."
    Write-Host "A line in the profile extensions tests for Win 7 then imports PSReadLine if required."
    if (-not (Get-Module -ListAvailable PSReadLine)) { 
        if ($PSver -gt 4) {
            Install-Module PSReadLine -Scope CurrentUser -Force -Verbose 
        }
        else {
            Write-Host "Need to be on PS v5 or higher to run 'Install-Module PSReadLine'" -ForegroundColor Red
        }
    }
    Import-Module PSReadLine -Scope Local -EA Silent
    $x = ""; foreach ($i in (get-command -module PSReadLine).Name) {$x = "$x, $i"} ; "" ; Write-Wrap $x.trimstart(", ") ; ""

    Write-Host ''
    Write-Host "Install-Module PSScriptAnalyzer (Script Analysis Tool)" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "   Get-ScriptAnalyzerRule, Invoke-Formatter, Invoke-ScriptAnalyzer"
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module PSScriptAnalyzer" -ForegroundColor Yellow
    # if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) { Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -Verbose }
    Install-ModuleToDirectory PSScriptAnalyzer $UserModulesPath

    Write-Host ''
    Write-Host "Install-Module Sudo" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "   New-SudoSession, Remove-SudoSession, Restore-OriginalSystemConfig, Start-SudoSession"
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module sudo" -ForegroundColor Yellow
    # if (-not (Get-Module -ListAvailable Sudo)) { Install-Module Sudo -Scope CurrentUser -Force -Verbose }
    Install-ModuleToDirectory Sudo $UserModulesPath

    Write-Host ''
    Write-Host "Install-Module Posh-Git (Git Management Cmdlets)" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "When you run Import-Module posh-git, posh-git checks to see if the PowerShell default prompt is the"
    Write-Host "current prompt. If it is, then posh-git will install a posh-git default prompt that looks like this in v0.x:"
    Write-Host "C:\Users\Keith\GitHub\posh-git [master ...> (the burger icon)"
    Write-Host "View details on `$GitPromptSettings here:"
    Write-Host "https://github.com/dahlbyk/posh-git/wiki/Customizing-Your-PowerShell-Prompt"
    Write-Host "Write-GitStatus, Write-Prompt, Write-VcsStatus"
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module posh-git" -ForegroundColor Yellow
    # if (-not (Get-Module -ListAvailable Posh-Git)) { Install-Module Posh-Git -Scope CurrentUser -Force -Verbose }
    Install-ModuleToDirectory Posh-Git $UserModulesPath

    Write-Host ''
    Write-Host "Install-Module Posh-Gist (Gist Management Cmdlets)" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "   Get-Gist, Get-GistCommits, Get-GistStar, New-Gist, Remove-Gist, Update-Gist"
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module posh-gist" -ForegroundColor Yellow
    # if (-not (Get-Module -ListAvailable Posh-Gist)) { Install-Module Posh-Gist -Scope CurrentUser -Force -Verbose }
    Install-ModuleToDirectory Posh-Gist $UserModulesPath

    # Write-Host ''
    # Write-Host "Install-Module PowerShellForGitHub   # (GitHub Management Cmdlets)" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "This is a PowerShell module that provides command-line interaction and automation for the GitHub v3 API."
    # Write-Host "https://github.com/microsoft/PowerShellForGitHub/blob/master/USAGE.md#examples"
    # Write-Host "http://stevenmaglio.blogspot.com/2019/08/powershellforgithubadding-get.html"
    # Write-Host "https://wilsonmar.github.io/powershell-github/"
    # Write-Host "https://hant.kutu66.com/GitHub/article_142903 (need to translate)"
    # Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module PowerShellForGitHub" -ForegroundColor Yellow
    # Install-ModuleToDirectory PowerShellForGitHub $ModulesPath

    Write-Host ''
    Write-Host "Install-Module Posh-SSH (SSH Cmdlets)" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "   Get-PoshSSHModVersion, Get-SFTPChildItem, Get-SFTPContent, Get-SFTPLocation, Get-SFTPPathAttribute,"
    # Write-Host "   Get-SFTPSession, Get-SSHPortForward, Get-SSHSession, Get-SSHTrustedHost, Invoke-SSHCommand,"
    # Write-Host "   Invoke-SSHCommandStream, Invoke-SSHStreamExpectAction, Invoke-SSHStreamExpectSecureAction,"
    # Write-Host "   Invoke-SSHStreamShellCommand, Move-SFTPItem, New-SFTPFileStream, New-SFTPItem, New-SFTPSymlink,"
    # Write-Host "   New-SSHDynamicPortForward, New-SSHLocalPortForward, New-SSHRemotePortForward, New-SSHShellStream,"
    # Write-Host "   New-SSHTrustedHost, Remove-SFTPItem, Remove-SFTPSession, Remove-SSHSession, Remove-SSHTrustedHost,"
    # Write-Host "   Rename-SFTPFile, Set-SFTPContent, Set-SFTPLocation, Set-SFTPPathAttribute, Start-SSHPortForward,"
    # Write-Host "   Stop-SSHPortForward, Test-SFTPPath, Get-SCPFile, Get-SCPFolder, Get-SCPItem, Get-SFTPFile,"
    # Write-Host "   Get-SFTPItem, New-SFTPSession, New-SSHSession, Set-SCPFile, Set-SCPFolder, Set-SCPItem, Set-SFTPFile,"
    # Write-Host "   Set-SFTPFolder, Set-SFTPItem"
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module posh-ssh" -ForegroundColor Yellow -NoNewline ; Write-Host "   # fimo *ssh* for other SSH tools" -ForegroundColor Green
    # if (-not (Get-Module -ListAvailable Posh-SSH)) { Install-Module Posh-SSH -Scope CurrentUser -Force -Verbose }
    # (gcm -mod posh-ssh | select Name | % { $_.Name + "," } | Out-String).replace("`r`n", " ").trim(", ")
    Install-ModuleToDirectory Posh-SSH $UserModulesPath

    Write-Host ''
    Write-Host "Install-Module PSColor (Color Get-ChildItem / gci / dir / ls output)" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "   Out-Default"
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module pscolor" -ForegroundColor Yellow
    Write-Host "Note: modifies Out-Default, so do not import by default, have setup 'color' function"
    Write-Host "in profile extensions to activate this when required."
    # if (-not (Get-Module -ListAvailable PSColor)) { Install-Module PSColor -Scope CurrentUser -Force -Verbose }
    Install-ModuleToDirectory PSColor $UserModulesPath

    Write-Host ''
    Write-Host "Install-Module Windows-ScreenFetch (System Utility)" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "View Module Contents: " -NoNewLine ; Write-Host "get-command -module windows-screenfetch" -ForegroundColor Yellow
    # if (-not (Get-Module -ListAvailable Windows-ScreenFetch)) { Install-Module Windows-ScreenFetch -Scope CurrentUser -Force -Verbose }
    Install-ModuleToDirectory Windows-ScreenFetch $UserModulesPath

    # Write-Host ''
    # Write-Host "Install-Module HistoryPx -AllowClobber (Enhanced Get-History tools)" -ForegroundColor Yellow -BackgroundColor Black
    # Write-Host "   Get-CaptureOutputConfiguration, Get-ExtendedHistoryConfiguration, Set-CaptureOutputConfiguration"
    # Write-Host "   Set-ExtendedHistoryConfiguration, Clear-History, Get-History, Out-Default"
    # Write-Host 'HistoryPx uses proxy commands to add extended history information to'
    # Write-Host 'PowerShell. This includes the duration of a command, a flag indicating whether'
    # Write-Host 'a command was successful or not, the output generated by a command (limited to'
    # Write-Host 'a configurable maximum value), the error generated by a command, and the'
    # Write-Host 'actual number of objects returned as output and as error records.  HistoryPx'
    # Write-Host 'also adds a "__" variable to PowerShell that captures the last output that you'
    # Write-Host 'may have wanted to capture, and includes commands to configure how it decides'
    # Write-Host 'when output should be captured.  Lastly, HistoryPx includes commands to manage'
    # Write-Host 'the memory footprint that is used by extended history information.'
    # Write-Host "View Module Contents: " -NoNewLine ; Write-Host "gcm -module historypx" -ForegroundColor Yellow
    # Write-Host "https://poshoholic.com/2014/10/30/making-history-more-fun-with-powershell/"
    # Write-Host "https://poshoholic.com/2014/10/21/transform-repetitive-script-blocks-into-invocable-snippets-with-snippetpx/"
    # Write-Host "https://poshoholic.com/2014/10/31/raise-your-powershell-game-with-historypx-debugpx-and-typepx/"
    # # if (-not (Get-Module -ListAvailable HistoryPx)) { Install-Module HistoryPx -AllowClobber -Scope CurrentUser -Force -Verbose }
    # Install-ModuleToDirectory HistoryPx $ModulesPath

    # Handling Subtitle files in PowerShell
    # https://github.com/KUTlime/PowerShell-Subtitle-Module
    # https://www.powershellgallery.com/packages/Subtitle/1.0.1.0
    # https://videoconverter.iskysoft.com/video-tips/download-subtitles.html

    Write-Host ''
    Write-Host "Example of querying Modules" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "Sort all available module sorting by version (which is 'Name' field in this view)"
    Write-Host "Get-Module -ListAvailable | group version | sort Name -Descending" -ForegroundColor Yellow
    #           Get-Module -ListAvailable | group version | sort Name -Descending | Out-Host
    Write-Host ''
    Write-Host "To show more info for those at a given version (say 2.0.0.0)"
    Write-Host "Get-Module -ListAvailable | ? { `$_.Version -eq '2.0.0.0' }" -ForegroundColor Yellow
    #           Get-Module -ListAvailable | ? { $_.version -eq '2.0.0.0' } | select Version,Name | sort -Descending | ft | Out-Host
    Write-Host "More info on commands in a specific module from this view (e.g. the VpnClient Module):"
    Write-Host "Get-Command -Module VpnClient | ft" -ForegroundColor Yellow
    #           Get-Command -Module VpnClient | ft | Out-Host
}




####################
#
# Note that WASP (Windows Automation Snap-in for Powershell) is fully available from PowerShell Gallery again:
#     Install-Module WASP
#
####################
# https://eddytechdotnet.wordpress.com/2016/02/24/using-wasp-and-powershell-for-powerful-windows-gui-automation/

# Also note (149 MB download!): https://archive.codeplex.com/?p=uiautomation    https://github.com/apetrovskiy/STUPS/tree/master/UIA

####################
#
# Very interesting IE automation of form entry (IE is terrible, but we still use it for Cognos at ING for example
#
####################
# https://cmdrkeene.com/automating-internet-explorer-with-powershell/



####################
#
# Key / Mouse detection, AutoHotkey functionality
#
####################

function Move-Mouse {
    # https://gist.github.com/MatthewSteeples/ce7114b4d3488fc49b6a
    # As a way to keep the system from going to sleep, this might work, might now, have to test
    # Can also use things like "Espresso" that sit in the tray and help with keep-alives

    Add-Type -AssemblyName System.Windows.Forms

    while ($true)
    {
        $p = [System.Windows.Forms.Cursor]::Position
        $x = ($p.X % 500) + 1
        $y = ($p.Y % 500) + 1
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
        Start-Sleep -Seconds 10
    }
}

function Template-CheckScheduledTasks {

    # This is testing for Scheduled Tasks and not Processes as shown by TaskManager

    # Method 1
    $taskName = "FireFoxMain"
    $taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -like $taskName }
    if($taskExists) {
        # Do whatever 
    } else {
        # Do whatever
    }

    # Method 2
    Get-ScheduledTask -TaskName "Task Name" -ErrorAction SilentlyContinue -OutVariable task

    # Method 3
    if (!$task) {
        # task does not exist, otherwise $task contains the task object
    }

    # Method 4
    if (Get-ScheduledTask FirefoxMaint -ErrorAction Ignore) { "found" } else { "not found" }
}

# This relates to my StackOverflow question about all module components not being applied 
# when the module is loaded; this forces all to be exported into the current session.
Export-ModuleMember -Alias * -Function *
