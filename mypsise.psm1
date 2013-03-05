function Set-PSISEGitHubTheme {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $fonts = New-Object System.Drawing.Text.InstalledFontCollection
    $font_fams = $fonts.Families
    $font_prefs = New-Object System.Collections.Specialized.OrderedDictionary
    $font_prefs.Add('Monaco', 10)
    $font_prefs.Add('DejaVu Sans Mono', 10)
    $font_prefs.Add('Lucida Console', 11)
    $font_prefs.Add('Consolas', 11)
    $font_prefs.Keys | % {
        $pref_font = $_
        if ($font_fams -contains $pref_font) {
            $psISE.Options.FontName = $pref_font
            $psISE.Options.FontSize = $font_prefs.Item($pref_font)
            break
        }
    }
    $psISE.Options.ScriptPaneBackgroundColor = '#141414'
    $psise.Options.TokenColors['Attribute'] = '#f8f8f8'
    $psise.Options.TokenColors['Command'] = '#dad085'
    $psise.Options.TokenColors['CommandArgument'] = '#f8f8f8' #Function name too.
    $psise.Options.TokenColors['CommandParameter'] = '#dad085'
    $psise.Options.TokenColors['Comment'] = '#5f5a60'
    $psise.Options.TokenColors['GroupEnd'] = '#f8f8f8'
    $psise.Options.TokenColors['GroupStart'] = '#f8f8f8'
    $psise.Options.TokenColors['Keyword'] = '#cda869'
    $psise.Options.TokenColors['LineContinuation'] = '#f8f8f8'
    $psise.Options.TokenColors['LoopLabel'] = '#f8f8f8'
    $psise.Options.TokenColors['Member'] = '#f8f8f8'
    $psise.Options.TokenColors['NewLine'] = '#cda869'
    $psise.Options.TokenColors['Number'] ='#cf6a4c'
    $psise.Options.TokenColors['Operator'] = '#cda869'
    $psise.Options.TokenColors['Position'] = '#f8f8f8'
    $psise.Options.TokenColors['StatementSeparator'] = '#f8f8f8'
    $psise.Options.TokenColors['String'] = '#8f9d6a'
    $psise.Options.TokenColors['Type'] = '#f8f8f8'
    $psise.Options.TokenColors['Unknown'] = '#f8f8f8'
    $psise.Options.TokenColors['Variable'] = '#7587a6'
}
