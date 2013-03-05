#region THEME
$is_IseV2 = ($host.Name -eq 'Windows PowerShell ISE Host' -and $host.Version.ToString() -eq '2.0')

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
$fonts = New-Object System.Drawing.Text.InstalledFontCollection
$font_fams = $fonts.Families
$font_prefs = New-Object System.Collections.Specialized.OrderedDictionary
if ($is_IseV2) {$font_bigger = 3} else {$font_bigger = 0}
$font_prefs.Add('Monaco', 10 + $font_bigger)
$font_prefs.Add('DejaVu Sans Mono', 10 + $font_bigger)
$font_prefs.Add('Lucida Console', 11 + $font_bigger)
$font_prefs.Add('Consolas', 11 + $font_bigger)
foreach ($pref_font in $font_prefs.Keys) {
    if ($font_fams -contains $pref_font) {
        $psISE.Options.FontName = $pref_font
        $psISE.Options.FontSize = $font_prefs.Item($pref_font)
        break
    }
}

if ($is_IseV2) {
    $psISE.Options.OutputPaneBackgroundColor = '#141414'
    $psISE.Options.OutputPaneTextBackgroundColor = '#141414'
    $psISE.Options.OutputPaneForegroundColor = '#f8f8f8'
    $psISE.Options.CommandPaneBackgroundColor = '#141414'
} else {
    $psise.Options.ConsolePaneBackgroundColor = '#141414'
    $psise.Options.ConsolePaneTextBackgroundColor = '#141414'
    $psise.Options.ConsolePaneForegroundColor = '#f8f8f8'
}

$psISE.Options.ScriptPaneBackgroundColor = '#141414'
$color_map = @{
    Attribute = '#f8f8f8'
    Command = '#dad085'
    CommandArgument = '#f8f8f8' #Function name too.
    CommandParameter = '#dad085'
    Comment = '#5f5a60'
    GroupEnd = '#f8f8f8'
    GroupStart = '#f8f8f8'
    Keyword = '#cda869'
    LineContinuation = '#f8f8f8'
    LoopLabel = '#f8f8f8'
    Member = '#f8f8f8'
    NewLine = '#cda869'
    Number = '#cf6a4c'
    Operator = '#cda869'
    Position = '#f8f8f8'
    StatementSeparator = '#f8f8f8'
    String = '#8f9d6a'
    Type = '#f8f8f8'
    Unknown = '#f8f8f8'
    Variable = '#7587a6'
}
$color_map2 = @{
    Attribute = '#cda869'
    CharacterData = '#7587a6'
    Comment	= '#5f5a60'
    CommentDelimiter = '#5f5a60'
    ElementName	= '#dad085'
    MarkupExtension	= '#8f9d6a'
    Quote = '#8f9d6a'
    QuotedString = '#8f9d6a'
    Tag	= '#f8f8f8'
    Text = '#8f9d6a'
}
foreach ($token_name in $color_map.Keys) {
    $color_value = $color_map.Item($token_name)
    $psise.Options.TokenColors[$token_name] = $color_value
    if (-not $is_IseV2) {
        $psise.Options.ConsoleTokenColors[$token_name] = $color_value
    }
}
if (-not $is_IseV2) {
    foreach ($token_name in $color_map2.Keys) {
        $color_value = $color_map2.Item($token_name)
        $psise.Options.XmlTokenColors[$token_name] = $color_value
    }
}
#endregion THEME