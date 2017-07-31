#!/usr/local/bin/powershell
#========================================
# NAME      : ConsoleTheme.psm1
# LANGUAGE  : PowerShell
# AUTHOR    : Bryan Dady
# Version   : 0.1.2
# UPDATED   : 07/30/2017
#========================================
[CmdletBinding(SupportsShouldProcess)]
param ()
# Set-StrictMode -Version latest

$Global:CTRootPath = Split-Path -Path $MyInvocation.MyCommand.Path -Resolve
Write-Verbose -Message "Dot-sourcing $(Join-Path -Path $Global:CTRootPath -ChildPath ConsoleTheme.ps1)"
. (Join-Path -Path $Global:CTRootPath -ChildPath ConsoleTheme.ps1)

Write-Verbose -Message 'Importing colors.csv to $colors'
New-Variable -Name colors -Value (Import-CSV -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'colors.csv' -Resolve)) -Description 'Custom table, populated via Import-CSV. Defined at scope at root of module, so it is shared across functions.' -Visibility Private -Scope 0 -Force
#$colors = Import-CSV -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'colors.csv' -Resolve)

Get-Variable -Name colors | Select-Object -Property Name,Value,scope -Verbose


Function Get-ColorCode {
    <#
        .SYNOPSIS
            Translates a color's common name to an 'RGB' color code
        .DESCRIPTION
            Using the colors.csv file, and accepting a color name string as a parameter to this function, it returns the matching 'RGB' color code 
        .EXAMPLE
            PS> Get-ColorCode White
            
            {65535, 65535, 65535}
            
        .INPUTS
            Inputs (if any)
        .OUTPUTS
            Output (if any)
        .NOTES
            General notes
    #>    
    [CmdletBinding()]
    param(
        [Alias('Color','Name')]
        [String]$ColorName = '*'
    )

    if (-not ($colors)) {
        Write-Verbose -Message 'Importing colors.csv to $colors'
        $colors = Import-CSV -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'colors.csv' -Resolve)
    }

    Write-Verbose -Message "Looking up `$ColorName $ColorName"
    return $($colors | Where-Object -FilterScript {$PSItem.Name -like "*$ColorName*"} | Select-Object -Property Name,RGB)

}

Function Get-ColorName {
    <#
        .SYNOPSIS
            Translates an 'RGB' color code to a color name
        .DESCRIPTION
            Using the colors.csv file, and accepting an 'RGB' color code as a parameter to this function, it returns the matching color name string
        .EXAMPLE
            PS> Get-ColorName {65535, 65535, 65535}
            
            White

        .INPUTS
            Inputs (if any)
        .OUTPUTS
            Output (if any)
        .NOTES
            General notes
    #>    
    [CmdletBinding()]
    param(
        [Alias('Code','Color')]
        [String]$ColorCode = '*'
    )

    if (-not ($colors)) {
        Write-Verbose -Message 'Importing colors.csv to $colors'
        $colors = Import-CSV -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'colors.csv' -Resolve)
    }

    Write-Verbose -Message "Looking up `$ColorCode $ColorCode"
    Write-Verbose -Message "return `$colors | Where-Object -FilterScript {`$PSItem.RGB -eq `"$ColorCode`"} | Select-Object -Property Name,RGB"
    $ColorName = $colors | Where-Object -FilterScript {$PSItem.RGB -eq $ColorCode} | Select-Object -Property Name,RGB
    
    Write-Debug -Message "`$ColorName is $ColorName"

    return $ColorName
}

Function Get-TerminalColor {
    <#
        .SYNOPSIS
            Gets the active color settings from the Terminal App
        .DESCRIPTION
            Enumerates the current color settings in the Terminal app
        
            In PowerShell Core (on non-Windows OS), the console / terminal emulation application that hosts the PowerShell 'ConsoleHost'
            Does not necessarily respect the color / theme preferences expressed in the $Host.UI.RawUI.BackgroundColor or $Host.UI.RawUI.ForegroundColor properties.

        .EXAMPLE
            PS> Get-TerminalColor

            Shows the current Terminal background color setting is Blue and the text color setting is White

        .NOTES
            To be added later
    #>
    [CmdletBinding()]
    param()
    <#
        tell application "Terminal"
            tell selected tab of front window
                get normal text color
                get background color
            end tell
        end tell
    #>

    Write-Debug -Message '& osascript (Join-Path -Path $Global:CTRootPath -ChildPath Get-ForegroundColor.applescript -Resolve)'
    Write-Verbose -Message "& osascript $(Join-Path -Path $Global:CTRootPath -ChildPath 'Get-ForegroundColor.applescript' -Resolve)"
    $script:FGColorCode = & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Get-ForegroundColor.applescript' -Resolve)
    $script:FGColorCode = "{$script:FGColorCode}"
    Write-Verbose -Message "`$script:FGColorCode is $script:FGColorCode"

    Write-Debug -Message '& osascript (Join-Path -Path $Global:CTRootPath -ChildPath Get-BackgroundColor.applescript -Resolve)'
    Write-Verbose -Message "& osascript $(Join-Path -Path $Global:CTRootPath -ChildPath 'Get-BackgroundColor.applescript' -Resolve)"
    $script:BGColorCode = & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Get-BackgroundColor.applescript' -Resolve)
    $script:BGColorCode = "{$script:BGColorCode}"
    Write-Verbose -Message "`$script:BGColorCode is $script:BGColorCode"

    Write-Verbose -Message "Get-ColorName -ColorCode $script:FGColorCode"
    $script:FGColorName = (Get-ColorName -ColorCode $script:FGColorCode | Select-Object -Property Name).Name

    Write-Verbose -Message " # `$script:FGColorName is $script:FGColorName"

    if ($script:FGColorName.Contains(',')) {
        #Split Name "array"
        $script:FGCollection = $script:FGColorName.Split(',')
    $script:FGCollection
        if ($script:FGCollection.Length -ge 2) {
            # Prefer the formatting of the last available Name string
            $script:FGColorName = $script:FGCollection[-1]
        } else {
            # play it safe and specify collection item 1
            $script:FGColorName = $script:FGCollection[0]
        }
    }

    Write-Verbose -Message "Get-ColorName -ColorCode $script:BGColorCode"
    $script:BGColorName = (Get-ColorName -ColorCode $script:BGColorCode | Select-Object -Property Name).Name

    Write-Verbose -Message " # `$script:BGColorName is $script:BGColorName"

    if ($script:BGColorName.Contains(',')) {
        #Split Name "array"
        $script:BGCollection = $script:BGColorName.Split(',')
        if ($script:BGCollection.Length -ge 2) {
            # Prefer the formatting of the last available Name string
            $script:BGColorName = $script:BGCollection[-1]
        } else {
            # play it safe and specify collection item 1
            $script:BGColorName = $script:BGCollection[0]
        }
    }

    #'Optimize New-Object invocation, based on Don Jones' recommendation: https://technet.microsoft.com/en-us/magazine/hh750381.aspx
    $Private:properties = [ordered]@{
        'ForegroundColorName' = $script:FGColorName
        'BackgroundColorName' = $script:BGColorName
        'ForegroundColorCode' = $script:FGColorCode
        'BackgroundColorCode' = $script:BGColorCode
    }
    $TerminalColor = New-Object -TypeName PSObject -Prop $properties
    
    return $TerminalColor
}

Function Get-Font {
    <#
        .SYNOPSIS
            Retrieves Font type (Name) and Font Size from the current Terminal.app window
        .DESCRIPTION
            Parses AppleScript results ...
        .EXAMPLE
            PS> Get-Font

            FontName        FontSize
            --------        --------
            Menlo-Regular   11

        .INPUTS
            Inputs (if any)
        .OUTPUTS
            Output (if any)
        .NOTES
            General notes
    #>    
    [CmdletBinding()]
    param()

    Write-Verbose -Message 'Get-Font.applescript'
    $FontName = & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Get-Font.applescript' -Resolve)
    Write-Verbose -Message 'Get-FontSize.applescript'
    $FontSize = & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Get-FontSize.applescript' -Resolve)
     
    #'Optimize New-Object invocation, based on Don Jones' recommendation: https://technet.microsoft.com/en-us/magazine/hh750381.aspx
    $Private:properties = [ordered]@{
        'FontName' = $FontName
        'FontSize' = $FontSize
    }
    $Font = New-Object -TypeName PSObject -Prop $properties
    
    return $Font
}

Function Set-Font {
    <#
        .SYNOPSIS
            Sets / updates the font type (name) and font size in the current Terminal.app window
        .DESCRIPTION
            Invokes AppleScript script files
        .EXAMPLE
            PS> Set-Font -FontName Consolas -FontSize 12

    #>    
    [CmdletBinding()]
    param(
        [Alias('Font','Name')]
        [String]$FontName
        ,
        [Alias('Size')]
        [String]$FontSize
    )

    Write-Verbose -Message "Set-Font -FontName $FontName -FontSize $FontSize"
    "tell application ""Terminal"" to set font of window 1 to ""$FontName""" | Out-File -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-Font.applescript')
    Write-Verbose -Message "$(Get-Content -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-Font.applescript'))"
    & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-Font.applescript')
    
    "tell application ""Terminal"" to set font size of window 1 to $FontSize" | Out-File -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-FontSize.applescript')
    Write-Verbose -Message "$(Get-Content -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-FontSize.applescript'))"
    & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-FontSize.applescript' -Resolve)
         
}
# Write-Verbose -Message 'Loading function Set-TerminalColor'
Function Set-TerminalColor {
    <#
        .SYNOPSIS
            Set / activate a new color option in the Terminal App
        .DESCRIPTION
            Select / express new color preferences to the Terminal app
        
            In PowerShell Core (on non-Windows OS), the console / terminal emulation application that hosts the PowerShell 'ConsoleHost'
            Does not necessarily respect the color / theme preferences expressed in the $Host.UI.RawUI.BackgroundColor or $Host.UI.RawUI.ForegroundColor properties

        .EXAMPLE
            PS> Set-TerminalColor -Background MidnightBlue

            Updates the Terminal background to a dark blue

        .EXAMPLE
            PS> Set-TerminalColor -Foreground SteelBlue

            Updates the Terminal background to a very light blue (gray)

            .EXAMPLE
            PS> Set-TerminalColor -BG Black -FG gray

            Updates the Terminal background to a very light blue (gray)


        .NOTES
            To be added later
    #>
    [CmdletBinding()]
    param(
        [Alias('BackgroundColor','Background')]
        [String]$BG
        ,
        [Alias('normal','Foreground','ForegroundColor','Text')]
        [String]$FG
    )

    <#
        & osascript `
        -e 'tell application "Terminal"' `
        -e 'tell selected tab of front window' `
        -e "set normal text color to $FG" `
        -e "set background color to $BG" `
        -e "end tell" `
        -e "end tell"
    #>

    if ($FG) {
        Write-Verbose -Message "Getting color code for $FG"
        Write-Debug -Message "Get-ColorCode -ColorName $FG"
        $FGColorMatch = Get-ColorCode -ColorName $FG
        if ($FGColorMatch.Length -ge 2) {
            # Prefer the first available Name match
            Write-Verbose -Message "Selecting color code for $FG match $($FGColorMatch[0].Name)"
            $FGColorCode = $FGColorMatch[0].RGB
        } else {
            $FGColorCode = $FGColorMatch.RGB
        }
        Write-Verbose -Message "`$FGColorCode is $FGColorCode"

        Write-Debug -Message "osascript -e tell application ""Terminal"" to set background color of window 1 to ${FGColorCode}"
        "tell application ""Terminal"" to set normal text color of window 1 to $FGColorCode" | Out-File -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-ForegroundColor.applescript')
        Write-Verbose -Message "$(Get-Content -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-ForegroundColor.applescript'))"
        & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-ForegroundColor.applescript')
    }

    if ($BG) {
        Write-Verbose -Message "Getting color code for $BG"
        Write-Debug -Message "Get-ColorCode -ColorName $BG"
        $BGColorMatch = Get-ColorCode -ColorName $BG
        if ($BGColorMatch.Length -ge 2) {
            # Prefer the first available Name match
            Write-Verbose -Message "Selecting color code for $BG match $($BGColorMatch[0].Name)"
            $BGColorCode = $BGColorMatch[0].RGB
        } else {
            $BGColorCode = $BGColorMatch.RGB
        }
        Write-Verbose -Message "`$BGColorCode is $BGColorCode"

        Write-Debug -Message "osascript -e tell application ""Terminal"" to set background color of window 1 to $BGColorCode"
        "tell application ""Terminal"" to set background color of window 1 to $BGColorCode" | Out-File -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-BackgroundColor.applescript')
        Write-Verbose -Message "$(Get-Content -Path (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-BackgroundColor.applescript'))"
        & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-BackgroundColor.applescript')
    }
}

Function Set-TerminalTitle {
    <#
        .SYNOPSIS
            Set Title of the current Terminal window
        .DESCRIPTION
            Set Title of the current Terminal window to PowerShell
        .EXAMPLE
            PS> Set-TerminalTitle
            
        .NOTES
            If/when I figure out how to properly format arguments from PowerShell, to osascript, or to the .applescript file,
            I'll update to handle passing the Title as a parameter.
    #>    
    [CmdletBinding()]
    param()
    Write-Verbose -Message "& osascript $(Join-Path -Path $Global:CTRootPath -ChildPath 'Set-TerminalTitle.applescript' -Resolve)"
    # tell application "Terminal" to set custom title of window 1 to "PowerShell"
    & osascript (Join-Path -Path $Global:CTRootPath -ChildPath 'Set-TerminalTitle.applescript' -Resolve)
} 
