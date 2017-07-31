#!/usr/local/bin/powershell
#========================================
# NAME      : ConsoleTheme.ps1
# LANGUAGE  : PowerShell
# AUTHOR    : Bryan Dady
# Version   : 0.1.2
# UPDATED   : 07/30/2017
#========================================
[CmdletBinding(SupportsShouldProcess)]
param ()
# Set-StrictMode -Version latest

#Region MyScriptInfo
    Write-Verbose -Message '[ConsoleTheme] Populating $MyScriptInfo'
    $script:MyCommandName = $MyInvocation.MyCommand.Name
    $script:MyCommandPath = $MyInvocation.MyCommand.Path
    $script:MyCommandType = $MyInvocation.MyCommand.CommandType
    $script:MyCommandModule = $MyInvocation.MyCommand.Module
    $script:MyModuleName = $MyInvocation.MyCommand.ModuleName
    $script:MyCommandParameters = $MyInvocation.MyCommand.Parameters
    $script:MyParameterSets = $MyInvocation.MyCommand.ParameterSets
    $script:MyRemotingCapability = $MyInvocation.MyCommand.RemotingCapability
    $script:MyVisibility = $MyInvocation.MyCommand.Visibility

    if (($null -eq $script:MyCommandName) -or ($null -eq $script:MyCommandPath)) {
        # We didn't get a successful command / script name or path from $MyInvocation, so check with CallStack
        Write-Verbose -Message "Getting PSCallStack [`$CallStack = Get-PSCallStack]"
        $CallStack = Get-PSCallStack | Select-Object -First 1
        # $CallStack | Select Position, ScriptName, Command | format-list # FunctionName, ScriptLineNumber, Arguments, Location
        $script:myScriptName = $CallStack.ScriptName
        $script:myCommand = $CallStack.Command
        Write-Verbose -Message "`$ScriptName: $script:myScriptName"
        Write-Verbose -Message "`$Command: $script:myCommand"
        Write-Verbose -Message 'Assigning previously null MyCommand variables with CallStack values'
        $script:MyCommandPath = $script:myScriptName
        $script:MyCommandName = $script:myCommand
    }

    #'Optimize New-Object invocation, based on Don Jones' recommendation: https://technet.microsoft.com/en-us/magazine/hh750381.aspx
    $Private:properties = [ordered]@{
        'CommandName'        = $script:MyCommandName
        'CommandPath'        = $script:MyCommandPath
        'CommandType'        = $script:MyCommandType
        'CommandModule'      = $script:MyCommandModule
        'ModuleName'         = $script:MyModuleName
        'CommandParameters'  = $script:MyCommandParameters.Keys
        'ParameterSets'      = $script:MyParameterSets
        'RemotingCapability' = $script:MyRemotingCapability
        'Visibility'         = $script:MyVisibility
    }
    $MyScriptInfo = New-Object -TypeName PSObject -Prop $properties
    
    Write-Verbose -Message '[ConsoleTheme] $MyScriptInfo populated'
    $MyScriptInfo

#End Region

if ((Get-Variable -Name IsOSX -ValueOnly -ErrorAction SilentlyContinue) -eq $true) { 
    Write-Verbose -Message 'Confirmed running on target OS: MacOS (OSX)'
} else {
    Write-Warning -Message 'This module is intended for MacOS (OSX)'
} 
