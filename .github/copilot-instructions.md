# Copilot Instructions

## General Guidelines
- Always resolve `ospp.vbs` from the script's folder using `$PSScriptRoot` or `$MyInvocation.MyCommand.Path` (script's directory) located at `C:\Users\hamlet23\Documents\OfficeDeploymentTool`, rather than from the user's Desktop; avoid hardcoded user paths. User prefers `ospp.vbs` to be located in the script folder and resolved using these methods instead of Desktop or hardcoded user paths. Update `cleanup-odt-licensing.ps1` to use script folder resolution.
- When constructing arrays of paths in PowerShell, avoid putting commas inside the `Join-Path` call. Use separate `Join-Path` calls as individual array elements (one per line) to prevent the parser from treating arguments as arrays.

