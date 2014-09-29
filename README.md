SSIS PowershellTools
====================

Tools for extracting information from SSIS files.

Get-SsisContent shows information on what components a SSIS file consists of, with type of Component, names, etc. It can e.g be used to find which DTS package uses a certain kind of component.

Get-SsisSql extracts information regarding SQL queries from a SSIS file. Database components are extracted as well as variables, as they can contain SQL queries too.

The Cmdlets are built to facilitate the powerfull Pipeline technique in Powershell. Files to be queried can be piped to the Cmdlets (e.g with Get-ChildItem) and the Cmdlets outputs the result in a pipeline friendly way, using custom types SsisTools.SsisContent and SsisTools.SsisSql. Intellisense is supported so you can see the properties when you pipe the result to other Cmdlets.

Usage:

Import-Module SsisPowerShellTools.psd1

Get-Help Get-SsisContent -Full
Get-Help Get-SsisSql -Full
