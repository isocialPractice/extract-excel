@echo off
::
:: ee-capture.cmd - capture `extract-excel --action:var` output into a cmd variable.
::
:: A child process cannot set variables in its parent shell, so extract-excel
:: emits a `set "NAME=value"` snippet. This helper runs the tool and executes
:: that snippet. Because it uses no `setlocal`, the assignment persists in the
:: caller's environment when the script is invoked with `call`.
::
:: Usage (call it, do not run in a new window):
::   call scripts\ee-capture.cmd book.xlsx --cell B2 --action:var _lastName
::   echo %_lastName%
::
:: Pass any extract-excel arguments; include an `--action:var <NAME>` so there
:: is an assignment to capture.

:: Force cmd-style `set` regardless of platform default.
set "EXTRACT_EXCEL_VAR_STYLE=cmd"

:: Locate the CLI: prefer a linked `extract-excel`, else the built dist.
where extract-excel >nul 2>nul && (
  set "_ee_cmd=extract-excel"
) || (
  set "_ee_cmd=node "%~dp0..\dist\cli.js""
)

:: Execute each emitted `set "..."` line in this (the caller's) scope.
for /f "usebackq delims=" %%L in (`%_ee_cmd% %*`) do %%L

set "_ee_cmd="
