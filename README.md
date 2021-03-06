# How to install Custom-Tools  
Installation does *not* require Administrator privileges and can run from a normal user console.  
  
`iex ((New-Object System.Net.WebClient).DownloadString('https://git.io/JqCtf'))`  

If there are any problems with installation, this is probably due to the recently more aggressive Windows Defender. This is not a bad thing, it's good for Microsoft to protect against viruses, but this toolkit is just a collection of functions and all of the function code is open on the Github project so there is nothing malicious in here.
`Set-MpPreference -DisableRealtimeMonitoring $true    # Disable`  
`... Now install Custon-Tools ...`  
`Set-MpPreference -DisableRealtimeMonitoring $false   # Enable`  


# Custom-Tools Basics  
• `BeginSystemConfig.ps1` controls setup of `ProfileExtensions.ps1` and `Custom-Tools.psm1`.  
• A single line is inserted into `$profile` that calls the ProfileExtensions (and the extensions script is placed in the profile folder). The ProfileExtensions holds only fundamental functions and definitions required to make changes to the `Custom-Tools` etc (to easily recover if something goes wrong with `Custom-Tools` etc).  
• `Custom-Tools.psm1` is installed under the users profile Module folder (see `$Env:PSModulePath`). Each of the functions in here is intended to be a completely stand-alone and portable so can extract anything to use for other projects.  
• To uninstall is easy simply remove the line from `$profile` that calls the extensions and run `Uninstall-Module Custom-Tools`.  
  
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
Note that the following may or may not be required as Windows Defender changes quite often. So try the above installer first and see if it works. Only temporarily stop Windows Defender if the above fails. This is sometimes required, as Windows Defender recently started heavily blocking many projects, even including big projects like Chocolatey (the main installer was completely blocked for a while unless you disabled Windows defender):  
https://github.com/chocolatey/choco/issues/2132  
https://theitbros.com/managing-windows-defender-using-powershell/  1
https://technoresult.com/how-to-disable-windows-defender-using-powershell-command-line/  
https://evotec.xyz/import-module-this-script-contains-malicious-content-and-has-been-blocked-by-your-antivirus-software/  
https://superuser.com/questions/1503345/i-disabled-real-time-monitoring-of-windows-defender-but-a-powershell-script-is  

If there are any Windows Defender issues, this can be bypassed by disabling Windows Defender briefly before installing (some people at Chocolatey thought that use of 'iex' / 'Invoke-Expression' or 'iwr' / 'Invoke-WebRequest' might have been the reason for this block, but I've not seen that confirmed) as follows:  
```
`# Temporarily stop Windows Defender (must be Administrator)`  
`# However, the above probably won't work as Windows Defender tries to protect itself and prevents this.`  
`sc.exe stop WinDefend   # Stop the Windows Defeder Service   # Stop the service`  
`... Now install Custon-Tools ...`  
`sc.exe start WinDefend   # Start the Windows Defeder Service  # Start the service`  
  
`# Disable RealTimeMonitoring (again, must be Administrator)`  
`Set-MpPreference -DisableRealtimeMonitoring $true    # Disable`  
`... Now install Custon-Tools ...`  
`Set-MpPreference -DisableRealtimeMonitoring $false   # Enable`  
```

# Notes

Used to remove all commit history (to reduce size of the `.git` folder): https://www.shellhacks.com/git-remove-all-commits-clear-git-history-local-remote/