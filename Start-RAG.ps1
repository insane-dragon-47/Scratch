param (
    [switch]$Shell
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$guiPath   = Join-Path $scriptDir "RAG-GUI.ps1"
$shellPath = Join-Path $scriptDir "RAG-Shell.ps1"

if (-not $env:OPENAI_API_KEY) {
    Write-Host "ERROR: Please set your OpenAI API key first using:"
    Write-Host '$env:OPENAI_API_KEY = "sk-..."'
    exit 1
}

if ($Shell) {
    & $shellPath
} else {
    & $guiPath
}