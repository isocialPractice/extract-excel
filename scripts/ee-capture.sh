#!/usr/bin/env bash
#
# ee-capture.sh — capture `extract-excel --action:var` output into shell variables.
#
# A child process cannot set variables in its parent shell, so extract-excel
# emits a sourceable `export NAME="value"` snippet instead. This helper runs the
# tool and evals that snippet in the *current* shell — so you must SOURCE it:
#
#   source ./scripts/ee-capture.sh book.xlsx --cell B2 --action:var _lastName
#   echo "$_lastName"          # => Anderson
#
# Any extract-excel arguments are accepted; include an `--action:var <NAME>`
# (or `--action:...,var var=<NAME> ...`) so there is an assignment to capture.

# Force POSIX-style `export` regardless of platform (e.g. bash on Windows).
export EXTRACT_EXCEL_VAR_STYLE=posix

# Resolve the extract-excel command: prefer a globally linked binary, else fall
# back to the built CLI relative to this script.
if command -v extract-excel >/dev/null 2>&1; then
  __ee_snippet="$(extract-excel "$@")"
else
  __ee_dir="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")/.." && pwd)"
  __ee_snippet="$(node "$__ee_dir/dist/cli.js" "$@")"
fi

# Evaluate the assignment(s) in the current shell.
eval "$__ee_snippet"
unset __ee_snippet __ee_dir
