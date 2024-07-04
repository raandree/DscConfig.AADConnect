@{
    PSDependOptions              = @{
        AddToPath  = $true
        Target     = 'output\RequiredModules'
        Parameters = @{
            Repository      = 'PSGallery'
            AllowPreRelease = $false
        }
    }

    InvokeBuild                  = 'latest'
    PSScriptAnalyzer             = 'latest'
    Pester                       = 'latest'
    Plaster                      = 'latest'
    ModuleBuilder                = 'latest'
    ChangelogManagement          = 'latest'
    Sampler                      = 'latest'
    'Sampler.GitHubTasks'        = 'latest'
    Datum                        = 'latest'
    'Datum.ProtectedData'        = 'latest'
    DscBuildHelpers              = 'latest'
    'DscResource.Test'           = 'latest'
    MarkdownLinkCheck            = 'latest'
    'DscResource.AnalyzerRules'  = 'latest'
    'DscResource.DocGenerator'   = 'latest'
    PSDesiredStateConfiguration  = 'latest'
    xDscResourceDesigner         = 'latest'

    #DSC Resources
    xPSDesiredStateConfiguration = '9.1.0'
    AADConnectDsc                = 'latest'

}
