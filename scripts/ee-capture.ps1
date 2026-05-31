<#
.SYNOPSIS
  Capture `extract-excel --action:var` output into PowerShell variables.

.DESCRIPTION
  A child process cannot set variables in its parent shell, so extract-excel
  emits a sourceable `$env:NAME = "value"` snippet. This helper runs the tool
  and invokes that snippet in the current session.

  Dot-source it so the assignment lands in your session:

    . .\scripts\ee-capture.ps1 book.xlsx --cell B2 --action:var _lastName
    $env:_lastName        # => Anderson

.NOTES
  Include an `--action:var <NAME>` (or `--action:...,var var=<NAME> ...`) among
  the arguments so there is an assignment to capture.
#>

# Force PowerShell-style assignment regardless of platform default.
$env:EXTRACT_EXCEL_VAR_STYLE = 'powershell'

# Resolve the CLI: prefer a linked `extract-excel`, else the built dist.
if (Get-Command extract-excel -ErrorAction SilentlyContinue) {
  $snippet = & extract-excel @args
} else {
  $cli = Join-Path (Split-Path $PSScriptRoot -Parent) 'dist\cli.js'
  $snippet = & node $cli @args
}

# Evaluate the emitted assignment(s) in the current session.
$snippet | ForEach-Object { if ($_ -ne '') { Invoke-Expression $_ } }
