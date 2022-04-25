# Custom-Tools Basics  
Tools and functions that integrate with the PowerShell console to make it more comfortable to use. There are no mass modifications to the environment, and the functions are all stand-alone so can be used as templates to build on or reuse in other scripts.  


# Custom-Tools Scripts  
All functions install and run under normal user space, there is no requirement for Administrator privileges. Full install is done from `BeginSystemConfig.ps1` which can be invoked directly by `iex ((New-Object System.Net.WebClient).DownloadString('https://git.io/JqCtf'))`  

`BeginSystemConfig.ps1` adds the `Custom-Tools.psm1` Module and adds two lines to the profile so that `ProfileExtensions.ps1` tools that are dot-sourced in all sessions. `ProfileExtensions` hold fundamental functions and definitions (such as the 'cd' alias which has been replaced by a superset of `Set-Location` to add more functionality. function for fast system navitation, recovery tools, and resizing functions for consoles `Custom-Tools` etc).  
• `Custom-Tools.psm1` is installed under the users profile Module folder (type `$Env:PSModulePath` to see that).  
• Each of the functions in here is deliberately stand-alone; there are no dependencies on each other so that each function can be completely portable for other projects as required.  
• To uninstall: simply remove the last two lines from `$profile` that calls the extensions and run `Uninstall-Module Custom-Tools` and the toolkit is completely uninstalled upon opening a new PowerShell console.  
  

# How to install Custom-Tools  

If there are any problems with installation, this is probably due to the recently more aggressive Windows Defender. It's good that Microsoft update Defender often, but this toolkit is just a collection of functions and all of the function code is open on the Github project so it is all safe.
`Set-MpPreference -DisableRealtimeMonitoring $true    # Disable`  
`# Now install Custon-Tools here ...`  
`Set-MpPreference -DisableRealtimeMonitoring $false   # Enable`  


**Simple Custom-Tools functions:**  
To view all installed Modules, use `mods`  
To drill in on a single module (such as Custom-Tools), use `mod custom-tools`  
To show just those functions that contain a given string, e.g. `mod custom-tools date`  
  
**Shorthand tools to explore functions in the Module:**  
To see the entire definition for a given function, use `def`, e.g. `def touch`  
The function `m` is used to provide access to all help functions *and* to the `about_` Topics.
• `m Para` - Show all `about_` Topics containing "Para".
• `m Foreach` - Show all `about_` Topics containing "Para".
• `m Foreach` - Show all `about_` Topics containing "Para".
Help works, but extensive help is not built into many Custom-Tools functions. e.g. `help def` and help can sometimes be awkward to navigate, so Custom-Tools containes some helpers to quickly get to Cmdlet and Function information:  
• `mm` (man module), quick way to get info on a module, so `mm pester`, but `mod pester` is more compact/efficient.
• `ms` (man syntax), can also use `syn`, so `syn def`, and `mparam` (man parameter) (can't use `mp` as that is a built-in alias), quick way to get info on a function parameter. These are very useful together to just get specific info on a command, e.g. To first see the syntax of a command, use `ms` and then to drill down and see detauled info on the `Filter` parameter:
`ms Get-ChildItem`  
`mparam Get-ChildItem Filter`  
• `me` (man examples), just show the examples for a given Cmdlet / Function, e.g. `me Get-ChildItem`  
• `mf` (man full), shows everything, full help page info, this is like Detailed, but expands every parameter property, e.g. `mf Get-ChildItem`  
  
Quick stuff: PowerShell variables, Environment Variables, useful system info:  
• `vars` / `getvars` : show currently defined PowerShell variables. https://stackoverflow.com/qu  
• `env` : Show Environment Variables (bit of a daft function, but I sometimes forget how to list environment variables)  
• `envgui` : Opens the Environment Variable GUI, `rundll32 sysdm.cpl,EditEnvironmentVariables`  
• `ver` : Show various version info. PowerShell, Windows, Office, etc  
• `sys` : Systems diagnostics, display a large amount of system information  

# Windows Defender  
Note that the following may or may not be required as Windows Defender changes quite often. So try the above installer first and see if it works. Only temporarily stop Windows Defender if the above fails. This is sometimes required, as Windows Defender recently started heavily blocking many projects, even including big projects like Chocolatey (the main installer was completely blocked for a while unless you disabled Windows Defender):  
https://github.com/chocolatey/choco/issues/2132  
https://theitbros.com/managing-windows-defender-using-powershell/  1
https://technoresult.com/how-to-disable-windows-defender-using-powershell-command-line/  
https://evotec.xyz/import-module-this-script-contains-malicious-content-and-has-been-blocked-by-your-antivirus-software/  
https://superuser.com/questions/1503345/i-disabled-real-time-monitoring-of-windows-defender-but-a-powershell-script-is  

If there are any Windows Defender issues, this can be bypassed by disabling Windows Defender briefly before installing (some people at Chocolatey thought that use of 'iex' / 'Invoke-Expression' or 'iwr' / 'Invoke-WebRequest' might have been the reason for this block, but I've not seen that confirmed) as follows:  
```
# Temporarily stop Windows Defender (must be Administrator)  
# However, the above probably won't work as Windows Defender tries to protect itself and prevents this.  
sc.exe stop WinDefend   # Stop the Windows Defeder Service   # Stop the service  

# Now install Custon-Tools here ...  
sc.exe start WinDefend   # Start the Windows Defeder Service  # Start the service  
  
# If the above did not work, try disabling RealTimeMonitoring ...
  
# Disable RealTimeMonitoring (again, must be Administrator)  
Set-MpPreference -DisableRealtimeMonitoring $true    # Disable  
  
# Now install Custon-Tools here...  
Set-MpPreference -DisableRealtimeMonitoring $false   # Enable  
```

# Notes

Used to remove all commit history (to reduce size of the `.git` folder): https://www.shellhacks.com/git-remove-all-commits-clear-git-history-local-remote/