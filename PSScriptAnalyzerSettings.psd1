@{
    # Copy this file into any folder with PowerShell scripts to define PSScriptAnalyzer rules

    ExcludeRules = @(
        'PSAvoidUsingWriteHost',                       # Suppress the rule that flags Write-Host
        'PSAvoidUsingCmdletAliases',                   # Suppress the rule that flags aliases
        'PSUseShouldProcessForStateChangingFunctions'  # Do not flag functions starting: New, Set, Remove, Start, Stop, Restart, Reset, Update
    )

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