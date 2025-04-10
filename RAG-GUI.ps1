Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module "$PSScriptRoot\RAG-Core.psm1" -Force

$contextDir = [Environment]::GetFolderPath("MyDocuments")
$model = $global:OpenAI_RAG_Model
$username = $env:USERNAME

$form = New-Object Windows.Forms.Form
$form.Text = "OpenAI RAG Assistant"
$form.Size = New-Object Drawing.Size(850, 600)
$form.BackColor = "Black"
$form.ForeColor = "White"
$form.Font = New-Object Drawing.Font("Segoe UI", 10)

$outputBox = New-Object Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.BackColor = "Black"
$outputBox.ForeColor = "White"
$outputBox.Dock = "Top"
$outputBox.Height = 320
$form.Controls.Add($outputBox)

$systemBox = New-Object Windows.Forms.TextBox
$systemBox.Multiline = $true
$systemBox.ScrollBars = "Vertical"
$systemBox.Text = "You are a helpful assistant. Use relevant document context if helpful."
$systemBox.BackColor = "Black"
$systemBox.ForeColor = "White"
$systemBox.Dock = "Top"
$systemBox.Height = 60
$form.Controls.Add($systemBox)

$inputBox = New-Object Windows.Forms.TextBox
$inputBox.Multiline = $true
$inputBox.ScrollBars = "Vertical"
$inputBox.BackColor = "DimGray"
$inputBox.ForeColor = "White"
$inputBox.Dock = "Bottom"
$inputBox.Height = 100
$form.Controls.Add($inputBox)

$submitButton = New-Object Windows.Forms.Button
$submitButton.Text = "Send"
$submitButton.Dock = "Bottom"
$form.Controls.Add($submitButton)

$reloadButton = New-Object Windows.Forms.Button
$reloadButton.Text = "Reindex Documents"
$reloadButton.Dock = "Bottom"
$form.Controls.Add($reloadButton)

$errorLogButton = New-Object Windows.Forms.Button
$errorLogButton.Text = "View Error Log"
$errorLogButton.Dock = "Bottom"
$form.Controls.Add($errorLogButton)

$chat = @()

$submitHandler = {
    $userInput = $inputBox.Text.Trim()
    if (-not $userInput) { return }

    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $outputBox.AppendText("[$username $timestamp]`r`n$userInput`r`n`r`n")

    $relevantChunks = Get-RelevantChunks -Directory $contextDir -Query $userInput -TopN 3
    $contextText = ($relevantChunks | ForEach-Object {
        "[from $($_.file)] $($_.content)"
    }) -join "`n---`n"

    if ($chat.Count -eq 0) {
        $chat += @{ role = "system"; content = $systemBox.Text }
    }

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

        $timestampAI = (Get-Date).ToString("HH:mm:ss")
        $outputBox.AppendText("[Assistant $timestampAI]`r`n$reply`r`n`r`n")

        $chat += @{ role = "assistant"; content = $reply }
        $inputBox.Clear()
    } catch {
        $outputBox.AppendText("ERROR: $_`r`n")
    }
}

$reloadHandler = {
    $outputBox.AppendText("[System $(Get-Date -Format 'HH:mm:ss')] Reindexing document context...`r`n")
    $null = Get-RelevantChunks -Directory $contextDir -Query "placeholder" -TopN 0
    $outputBox.AppendText("Done.`r`n`r`n")
}

$errorLogHandler = {
    $logPath = $global:OpenAI_RAG_ErrorLog
    $logText = if (Test-Path $logPath) { Get-Content $logPath -Raw } else { "No errors logged." }

    $logForm = New-Object Windows.Forms.Form
    $logForm.Text = "Unreadable Files"
    $logForm.Size = New-Object Drawing.Size(600, 400)
    $logForm.BackColor = "Black"
    $logForm.ForeColor = "White"
    $logForm.Font = New-Object Drawing.Font("Consolas", 9)

    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ReadOnly = $true
    $textBox.Dock = "Fill"
    $textBox.BackColor = "Black"
    $textBox.ForeColor = "White"
    $textBox.ScrollBars = "Vertical"
    $textBox.Text = $logText

    $logForm.Controls.Add($textBox)
    $logForm.ShowDialog()
}

$submitButton.Add_Click($submitHandler)
$reloadButton.Add_Click($reloadHandler)
$errorLogButton.Add_Click($errorLogHandler)

$inputBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter" -and $_.Modifiers -eq "Control") {
        $submitHandler.Invoke()
        $_.SuppressKeyPress = $true
    }
})

[void]$form.ShowDialog()