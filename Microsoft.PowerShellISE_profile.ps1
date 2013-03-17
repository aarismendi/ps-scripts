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
$psISE.Options.ScriptPaneForegroundColor = '#f8f8f8'
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

function Duplicate-Line
{
    $editor = $psISE.CurrentFile.Editor
    $caret_row = $editor.CaretLine
    $caret_col = $editor.CaretColumn
    $this_line_text = $editor.Text.Split("`n")[$caret_row - 1].TrimEnd([environment]::NewLine)
    $editor.SetCaretPosition($caret_row, $this_line_text.length + 1)
    $editor.InsertText("`r`n" + $this_line_text)    
    $editor.SetCaretPosition($caret_row, $caret_col)
}

#requires -version 2.0
## ISE-Comments module v 1.1
##############################################################################################################
## Provides Comment cmdlets for working with ISE
## ConvertTo-BlockComment - Comments out selected text with <# before and #> after
## ConvertTo-BlockUncomment - Removes <# before and #> after selected text
## ConvertTo-Comment - Comments out selected text with a leeding # on every line 
## ConvertTo-Uncomment - Removes leeding # on every line of selected text
##
## Usage within ISE or Microsoft.PowershellISE_profile.ps1:
## Import-Module ISE-Comments.psm1
##
## Note: The IsePack, a part of the PowerShellPack, also contains a "Toggle Comments" command,
##       but it does not support Block Comments
##       http://code.msdn.microsoft.com/PowerShellPack
##
##############################################################################################################
## History:
## 1.1 - Minor alterations to work with PowerShell 2.0 RTM and Documentation updates (Hardwick)
## 1.0 - Initial release (Poetter)
##############################################################################################################


## ConvertTo-BlockComment
##############################################################################################################
## Comments out selected text with <# before and #> after
## This code was originaly designed by Jeffrey Snover and was taken from the Windows PowerShell Blog.
## The original function was named ConvertTo-Comment but as it comments out a block I renamed it.
##############################################################################################################
function ConvertTo-BlockComment
{
    $editor = $psISE.CurrentFile.Editor
    $CommentedText = "<#`n" + $editor.SelectedText + "#>"
    # INSERTING overwrites the SELECTED text
    $editor.InsertText($CommentedText)
}

## ConvertTo-BlockUncomment
##############################################################################################################
## Removes <# before and #> after selected text
##############################################################################################################
function ConvertTo-BlockUncomment
{
    $editor = $psISE.CurrentFile.Editor
    $CommentedText = $editor.SelectedText -replace ("^<#`n", "")
    $CommentedText = $CommentedText -replace ("#>$", "")
    # INSERTING overwrites the SELECTED text
    $editor.InsertText($CommentedText)
}

## ConvertTo-Comment
##############################################################################################################
## Comments out selected text with a leeding # on every line
##############################################################################################################
function ConvertTo-Comment
{
    $editor = $psISE.CurrentFile.Editor
    $CommentedText = $editor.SelectedText.Split("`n")
    # INSERTING overwrites the SELECTED text
    $editor.InsertText( "#" + ( [String]::Join("`n#", $CommentedText)))
}

## ConvertTo-Uncomment
##############################################################################################################
## Comments out selected text with <# before and #> after
##############################################################################################################
function ConvertTo-Uncomment
{
    $editor = $psISE.CurrentFile.Editor
    $CommentedText = $editor.SelectedText.Split("`n") -replace ( "^#", "" )
    # INSERTING overwrites the SELECTED text
    $editor.InsertText( [String]::Join("`n", $CommentedText))
}

##############################################################################################################
## Inserts a submenu Comments to ISE's Custum Menu
## Inserts command Block Comment Selected to submenu Comments
## Inserts command Block Uncomment Selected to submenu Comments
## Inserts command Comment Selected to submenu Comments
## Inserts command Uncomment Selected to submenu Comments
##############################################################################################################
if (-not( $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus | where { $_.DisplayName -eq "Comments" } ) )
{
	$commentsMenu = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("_Comments",$null,$null) 
	$null = $commentsMenu.Submenus.Add("Block Comment Selected", {ConvertTo-BlockComment}, "Ctrl+SHIFT+B")
	$null = $commentsMenu.Submenus.Add("Block Uncomment Selected", {ConvertTo-BlockUncomment}, "Ctrl+Alt+B")
	$null = $commentsMenu.Submenus.Add("Comment Selected", {ConvertTo-Comment}, "Ctrl+SHIFT+C")
	$null = $commentsMenu.Submenus.Add("Uncomment Selected", {ConvertTo-Uncomment}, "Ctrl+Alt+C")
    $null = $commentsMenu.Submenus.Add("Duplicate Line", {Duplicate-Line}, "Ctrl+Alt+D")
}
