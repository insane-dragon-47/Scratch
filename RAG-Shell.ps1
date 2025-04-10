Import-Module "$PSScriptRoot\RAG-Core.psm1" -Force

$contextDir = [Environment]::GetFolderPath("MyDocuments")
$model = $global:OpenAI_RAG_Model
$username = $env:USERNAME

$chat = @(
    @{ role = "system"; content = "You are a helpful assistant. Use provided documents to improve your answers." }
)

Write-Host "=== OpenAI RAG Shell ==="
Write-Host "Context: $contextDir"
Write-Host "Type 'exit' to quit, 'reload' to reindex documents."
Write-Host "Enter multi-line messages. Press Ctrl+Z then Enter (Windows) or Ctrl+D (Linux) to submit.`n"

function Read-Multiline {
    Write-Host "`n[$username $(Get-Date -Format 'HH:mm:ss')]"
    $buffer = ""
    while ($true) {
        $line = [Console]::In.ReadLine()
        if ($line -eq $null) { break }
        $buffer += "$line`n"
    }
    return $buffer.Trim()
}

while ($true) {
    $userInput = Read-Multiline

    if (-not $userInput) { continue }
    if ($userInput -eq "exit") { break }
    if ($userInput -eq "reload") {
        Write-Host "`n[System $(Get-Date -Format 'HH:mm:ss')] Reindexing document context...`n" -ForegroundColor DarkGray
        continue
    }

    $relevantChunks = Get-RelevantChunks -Directory $contextDir -Query $userInput -TopN 3
    $contextText = ($relevantChunks | ForEach-Object {
        "[from $($_.file)] $($_.content)"
    }) -join "`n---`n"

    $chat += @{ role = "system"; content = "Relevant documents:`n$contextText" }
    $chat += @{ role = "user"; content = $userInput }

    $body = @{
        model = $model
        messages = $chat
    } | ConvertTo-Json -Depth 10

    $headers = @{
        "Authorization" = "Bearer $env:OPENAI_API_KEY"
        "Content-Type"  = "application/json"
    }

    try {
        $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Headers $headers -Method Post -Body $body
        $reply = $response.choices[0].message.content

        Write-Host "`n[Assistant $(Get-Date -Format 'HH:mm:ss')]" -ForegroundColor Yellow
        Write-Host "$reply`n" -ForegroundColor Yellow

        $chat += @{ role = "assistant"; content = $reply }
    } catch {
        Write-Host "ERROR: $_" -ForegroundColor Red
    }

    Write-Host "`n---"
}