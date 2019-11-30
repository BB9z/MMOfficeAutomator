#! /bin/sh
set -euo pipefail
cd "$(dirname "$0")"
pwsh -File "$PWD/Scripts/main.ps1"
