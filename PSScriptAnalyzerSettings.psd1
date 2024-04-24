# Can either create this as a global PSScriptAnalyzerSettings.psd1 file in the default PSScriptAnalyzer folder:
#   $env:UserProfile\Documents\WindowsPowerShell\Modules\PSScriptAnalyzer\1.19.1\PSScriptAnalyzer.psd1
# Or, can copy this file into any folder with PowerShell scripts to define PSScriptAnalyzer rules. 
# Run PSScriptAnalyzer directly via the command palette (Ctrl+Shift+P) "PSScriptAnalyzer: Analyze Script" or by right-clicking in the editor and selecting "PSScriptAnalyzer: Analyze Script".
# View Issues: After running the analyzer, any detected issues will be displayed in the "Problems" panel at the bottom of the editor. You can click on each issue to navigate to the corresponding line in the script.
# Navigate Between Issues: To navigate between the detected issues, you can use the keyboard shortcut F8 (by default). Pressing F8 will cycle through the issues one by one, highlighting each one in the editor and moving to the corresponding line.
# https://github.com/PowerShell/PSScriptAnalyzer/blob/master/RuleDocumentation/AvoidUsingCmdletAliases.md
# https://github.com/PowerShell/PSScriptAnalyzer/issues/214

@{

    # Each rule statement should be separated by ; and ExcludeRules by ,
    Rules = @{
        PSAvoidUsingCmdletNameExceptions = @{ AllowList = @('Help') };       # I use Help-xxx Funtions that I would like to keep
        # PSAvoidUsingCmdletAliases = @{ AllowList = @("ls", "cd", "gci") }    # Can allow specific rules here (comment this out if ignoring *all* aliases)
    }

    # Do not analyze the rules in ExcludeRules. Use it when you have commented out the IncludeRules section, which means that all
    # default rules will be included, *except* for those in the section below. Note that if a rule is in both IncludeRules and
    # ExcludeRules, the rule will be exluded.
    ExcludeRules = @(
        'PSAvoidUsingWriteHost',                       # Suppress the rule that flags Write-Host
        'PSAvoidUsingCmdletAliases',                   # Suppress the rule that flags aliases
        'PSUseShouldProcessForStateChangingFunctions'  # Do not flag functions starting: New, Set, Remove, Start, Stop, Restart, Reset, Update
    )
    # You can also use rule configuration to configure independent rules (for those rules that support this granularity):
    # Rules = @{ PSAvoidUsingCmdletAliases = @{ AllowList = @("ls", "cd", "gci") } }


    # IncludeRules = @()


    # Custom Rules: Specify a path to a folder containing custom rule modules. This allows you to extend the set of rules provided by default with custom rules tailored to your specific requirements.
    # CustomRulePath = @(
    #     'C:\Path\To\CustomRules'
    # )

    # Rule Severity: Define severity levels for each rule. Severity levels include Error, Warning, and Informational.
    # RuleSeverity = @{
    #     PSAvoidUsingWriteHost = 'Warning'
    #     PSUseShouldProcessForStateChangingFunctions = 'Informational'
    # }

    # Rule Configuration: Some rules allow for additional configuration options. You can define these options using the RuleConfiguration setting.
    # RuleConfiguration = @{
    #     PSAvoidUsingPlainTextForPassword = @{
    #         DescriptiveName = 'Check for plain text passwords'
    #         Enabled = $true
    #         ErrorMessage = 'Do not store passwords in plain text'
    #     }
    # }

    # DefaultRulesPath: Specifies the path to the folder containing built-in rule modules. This setting is typically predefined and doesn't need manual configuration unless you're customizing the location of the rule modules.
    # DefaultRulesPath = 'C:\Program Files\WindowsPowerShell\Modules\PSScriptAnalyzer\Rules'

    # RecurseCustomRules: Determines whether to recursively search for custom rules in subdirectories of the path specified in CustomRulePath.
    # RecurseCustomRules = $true

    # MinimumVersion: Specifies the minimum version of PSScriptAnalyzer required to process the configuration file.
    # MinimumVersion = '1.19.0'

}