# Custom-External.psm1
# Module containing various dynamically loaded tools and techniques
# from the internet.
# Only use older/stable/reliable repositories in here as there is a small danger associated
# with loading a script that has been compromised. However, the repositories below are all 
# very stable.


# https://github.com/janikvonrotz/awesome-powershell
# Awesome PowerShell Awesome Quality Assurance
# A curated list of delightful PowerShell packages and resources.
# PowerShell is a cross-platform (Windows, Linux, and macOS) automation and configuration tool
# that is optimized for dealing with structured data (e.g. JSON, CSV, XML, etc.), REST APIs,
# and object models. It includes a command-line shell and an associated scripting language.

# Very useful repositories to look at:
# https://github.com/proxb?tab=repositories
#    https://github.com/proxb/PowerShellModulesCentral
#    https://github.com/proxb/PowerShell
# https://github.com/janikvonrotz?tab=repositories
#    https://github.com/janikvonrotz/gistbox   Various PowerShell scripts
#    https://github.com/janikvonrotz/dotfiles
#    https://github.com/janikvonrotz/emoji-cheat-sheet

#== Windows 10 DeCrapifier and DeBloater ==
# https://www.makeuseof.com/windows-10-decrapifier-debloater/
# https://www.lifewire.com/pc-decrapifier-review-2626193
# DeCrapifier: https://community.spiceworks.com/scripts/show/4378-windows-10-decrapifier-18xx-19xx-2xxx
# DeBloater:   https://freetimetech.com/windows-10-clean-up-debloat-tool-by-ftt/
# https://github.com/eccko/TrashEraser-Debloat-Windows-10  

# function Enable-OnlineFunction {
# 
#     # using statements must be at the top of a script, i.e. cannot be inside a function
#     # using namespace System.Management.Automation.Language
#     # If this is not possible, need to inject "System.Management.Automation.Language" before Parser and FunctionDefinitionAst
#     # https://stackoverflow.com/questions/72069554/powershell-auto-load-functions-from-internet-on-demand
#     # Load-Function https://raw.githubusercontent.com/proxb/AsyncFunctions/master/Test-ConnectionAsync.ps1
#     # Ping-Subnet            # => now is available in your current session.
#     # Test-ConnectionAsync   # => now is available in your current session.
#     
#     [cmdletbinding()]
#     param(
#         [parameter(Mandatory, ValueFromPipeline)]
#         [uri] $URI
#     )
# 
#     process {
#         try {
#             $funcs = Invoke-RestMethod $URI
#             $ast = [System.Management.Automation.Language.Parser]::ParseInput($funcs, [ref] $null, [ref] $null)
#             foreach($func in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {   # [FuntionDefinitionAst]
#                 if($func.Name -in (Get-Command -CommandType Function).Name) {
#                     Write-Warning "$($func.Name) is already loaded! Skipping"
#                     continue
#                 }
#                 New-Item -Name "script:$($func.Name)" -Path function: -Value $func.Body.GetScriptBlock()
#             }
#         }
#         catch {
#             Write-Warning $_.Exception.Message
#         }
#     }
# }
# 
# functions Enable-PingDNSTools {
#     Enable-OnlineFunction https://raw.githubusercontent.com/proxb/AsyncFunctions/master/Test-ConnectionAsync.ps1
#     echo "'Ping-Subnet' is now available in your current session."
#     echo "'Test-ConnectionAsync'  is now available in your current session."
# }

# Import Online Functions
# Boe Prox https://github.com/proxb/

# function Invoke-RegExHelper {
# <#
# Invoke-RegExHelper
####...SYNOPSIS 
# A UI to help with writing Regular Expressions.
# https://github.com/proxb/RegExHelper
# This is a UI built using PowerShell and WPF that allows for simple Regular Expression checking by displaying the results in real time.
# 
# Currently this only supports a string match but future versions will allow for locating patterns in a log file or similiar groups of text.
# 
# Feedback and improvements are always welcome! Be sure to check out the Dev branch to help out with the log file regular expression helper.
# 
# You need to dot source the script to load the Invoke-RegExHelper function.
# #>
#     $toImport = irm "https://raw.githubusercontent.com/proxb/AsyncFunctions/master/Test-ConnectionAsync.ps1"
#     New-Module ([ScriptBlock]::Create($toImport)) | Out-Null
#     $MyInvocation.Line | iex
# }

function Test-ConnectionAsync {
    # https://github.com/proxb/AsyncFunctions
    $toImport = irm "https://raw.githubusercontent.com/proxb/AsyncFunctions/master/Test-ConnectionAsync.ps1"
    New-Module ([ScriptBlock]::Create($toImport)) | Out-Null
    $MyInvocation.Line | iex
}

function Ping-Subnet {
    # https://github.com/proxb/AsyncFunctions
    $toImport = (irm "https://raw.githubusercontent.com/proxb/AsyncFunctions/master/Test-ConnectionAsync.ps1").Replace([Text.Encoding]::UTF8.GetString((239,187,191)),"")
    New-Module ([ScriptBlock]::Create($toImport)) | Out-Null
    $MyInvocation.Line | iex
}

function Get-DNSHostEntryAsync {
    # https://github.com/proxb/AsyncFunctions
    $toImport = (irm "https://raw.githubusercontent.com/proxb/AsyncFunctions/master/Get-DNSHostEntryAsync.ps1").Replace([Text.Encoding]::UTF8.GetString((239,187,191)),"")
    New-Module ([ScriptBlock]::Create($toImport)) | Out-Null
    $MyInvocation.Line | iex
}

function Test-Port {
    # https://github.com/proxb/PowerShell_Scripts
    $toImport = (irm "https://raw.githubusercontent.com/proxb/PowerShell_Scripts/master/Test-Port.ps1").Replace([Text.Encoding]::UTF8.GetString((239,187,191)),"")
    New-Module ([ScriptBlock]::Create($toImport)) | Out-Null
    $MyInvocation.Line | iex
}

function Get-TCPResponse {
    # https://github.com/proxb/PowerShell_Scripts
    $toImport = (irm "https://raw.githubusercontent.com/proxb/PowerShell_Scripts/master/Get-TCPResponse.ps1").Replace([Text.Encoding]::UTF8.GetString((239,187,191)),"")
    New-Module ([ScriptBlock]::Create($toImport)) | Out-Null
    $MyInvocation.Line | iex
}

function Invoke-RegExHelper {
    # https://github.com/proxb/PowerShell_Scripts
    $toImport = (irm "https://raw.githubusercontent.com/proxb/RegExHelper/master/Invoke-RegExHelper.ps1").Replace([Text.Encoding]::UTF8.GetString((239,187,191)),"")
    New-Module ([ScriptBlock]::Create($toImport)) | Out-Null
    $MyInvocation.Line | iex
}

function Enable-PinnedItemsModule {
    # https://github.com/proxb/PinnedItem
    # Can only run as Admin
    Install-Module -Name PinnedItem
    # Install-Module -Name PinnedItem   # Dealing with Start Menu and Taskbar Pinned Items
    # Get-PinnedItem
    # New-PinnedItem -TargetPath "C:\Program Files (x86)\Internet Explorer\iexplore.exe" -Type TaskBar
    # $TargetPath = 'PowerShell.exe'
    # $ShortCutPath = 'WinDbg.lnk'
    # $Argument = "-ExecutionPolicy Bypass -NoProfile -NoLogo -Command `"& 'C:\users\proxb\desktop\Windbg.exe'`""
    # $Icon = 'C:\users\proxb\desktop\Windbg.exe'
    # New-PinnedItem -TargetPath $TargetPath -ShortCutPath $ShortcutPath -Argument $Argument -Type TaskBar -IconLocation $Icon
    # Get-PinnedItem -Type StartMenu | Where {$_.Name -eq 'Snipping Tool'} | Remove-PinnedItem 
}

function Install-ChrisTitusTools {
    # https://christitus.com/debloat-windows-10-2020/
    . iex ((New-Object System.Net.WebClient).DownloadString('https://git.io/JJ8R4'))
}

function RickASCII {
    # Rick Astley
    iex (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")
}
