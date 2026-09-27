@{
    # PSScriptAnalyzer-Einstellungen für easyONBOARDING 3.
    # Aufruf: Invoke-ScriptAnalyzer -Path . -Recurse -Settings ./PSScriptAnalyzerSettings.psd1
    # CI-Gate: Scripts/Invoke-EobQualityCheck.ps1 (bricht bei jedem Befund ab).
    #
    # Ausnahmen werden nicht global abgeschaltet, sondern direkt im Code per
    # [Diagnostics.CodeAnalysis.SuppressMessageAttribute(..., Justification = '...')] begründet.

    Severity            = @('Error', 'Warning', 'Information')
    IncludeDefaultRules = $true

    ExcludeRules        = @(
        # Die Regel wertet 'return @(...)' als Rückgabetyp Object[], obwohl PowerShell die Elemente
        # einzeln in die Pipeline gibt. [OutputType()] dokumentiert in diesem Projekt den Elementtyp.
        'PSUseOutputTypeCorrectly'
    )

    # PSUseCorrectCasing ist bewusst nicht aktiviert: In Kombination mit den übrigen Regeln führt sie
    # mit PSScriptAnalyzer 1.23/1.24 unter PowerShell 7.4 zu sporadischen internen Fehlern
    # ("The term 'Get-Command' is not recognized") und damit zu einem unzuverlässigen CI-Gate.

    Rules               = @{
        PSUseCompatibleSyntax                     = @{
            Enable         = $true
            TargetVersions = @('7.0')
        }
        PSPlaceOpenBrace                          = @{
            Enable             = $true
            OnSameLine         = $true
            NewLineAfter       = $true
            IgnoreOneLineBlock = $true
        }
        PSPlaceCloseBrace                         = @{
            Enable             = $true
            NewLineAfter       = $true
            IgnoreOneLineBlock = $true
            NoEmptyLineBefore  = $false
        }
        PSUseConsistentIndentation                = @{
            Enable              = $true
            IndentationSize     = 4
            Kind                = 'space'
            PipelineIndentation = 'IncreaseIndentationForFirstPipeline'
        }
        PSUseConsistentWhitespace                 = @{
            Enable                                  = $true
            CheckInnerBrace                         = $true
            CheckOpenBrace                          = $true
            CheckOpenParen                          = $true
            CheckOperator                           = $true
            CheckPipe                               = $true
            CheckPipeForRedundantWhitespace         = $true
            CheckSeparator                          = $true
            CheckParameter                          = $true
            # Ausgerichtete Zuweisungen in mehrzeiligen Hashtables sind erlaubt.
            IgnoreAssignmentOperatorInsideHashTable = $true
        }
        PSAvoidUsingDoubleQuotesForConstantString = @{
            Enable = $true
        }
        PSAvoidSemicolonsAsLineTerminators        = @{
            Enable = $true
        }
        PSAvoidExclaimOperator                    = @{
            Enable = $true
        }
    }
}
