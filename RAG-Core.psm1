$global:OpenAI_RAG_AppData = Join-Path $env:APPDATA "OpenAI-RAG"
$global:OpenAI_RAG_CachePath = Join-Path $OpenAI_RAG_AppData "embeddings.json"
$global:OpenAI_RAG_ErrorLog = Join-Path $OpenAI_RAG_AppData "unreadable-files.log"
$global:OpenAI_RAG_Model = "gpt-4"
$global:OpenAI_RAG_EmbeddingModel = "text-embedding-ada-002"

if (-not (Test-Path $OpenAI_RAG_AppData)) {
    New-Item -ItemType Directory -Path $OpenAI_RAG_AppData | Out-Null
}
if (-not (Test-Path $OpenAI_RAG_CachePath)) {
    @{} | ConvertTo-Json | Set-Content $OpenAI_RAG_CachePath
}

function Log-UnreadableFile {
    param([string]$Path, [string]$Reason)
    "$((Get-Date).ToString("s")) | $Path | $Reason" | Out-File -Append -Encoding UTF8 $global:OpenAI_RAG_ErrorLog
}

function Get-PdfText {
    param([string]$Path)
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $shell = New-Object -ComObject shell.application
        $folder = Split-Path $Path
        $file = Split-Path $Path -Leaf
        $folderItem = $shell.Namespace($folder).ParseName($file)
        $folderItem.InvokeVerb("Copy")
        Start-Sleep -Milliseconds 500
        return [Windows.Forms.Clipboard]::GetText()
    } catch {
        Log-UnreadableFile -Path $Path -Reason "Failed to read PDF: $_"
        return ""
    }
}

function Get-WordText {
    param([string]$Path)
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($Path)
        $text = $doc.Content.Text
        $doc.Close()
        $word.Quit()
        return $text
    } catch {
        Log-UnreadableFile -Path $Path -Reason "Failed to read Word document: $_"
        return ""
    }
}

function Get-ExcelText {
    param([string]$Path)
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $book = $excel.Workbooks.Open($Path)
        $sheet = $book.Sheets.Item(1)
        $range = $sheet.UsedRange
        $content = $range.Value2 | ForEach-Object { $_ -join " " } | Out-String
        $book.Close()
        $excel.Quit()
        return $content
    } catch {
        Log-UnreadableFile -Path $Path -Reason "Failed to read Excel document: $_"
        return ""
    }
}

function Get-PowerPointText {
    param([string]$Path)
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = $false
        $presentation = $ppt.Presentations.Open($Path, $true, $false, $false)
        $content = ""
        foreach ($slide in $presentation.Slides) {
            foreach ($shape in $slide.Shapes) {
                if ($shape.TextFrame -and $shape.TextFrame.HasText) {
                    $content += $shape.TextFrame.TextRange.Text + "`n"
                }
            }
        }
        $presentation.Close()
        $ppt.Quit()
        return $content
    } catch {
        Log-UnreadableFile -Path $Path -Reason "Failed to read PowerPoint: $_"
        return ""
    }
}

function Get-TextFromFile {
    param([string]$Path)
    try {
        switch -Regex ($Path) {
            '\.pdf$'   { return Get-PdfText -Path $Path }
            '\.txt$'   { return Get-Content $Path -Raw }
            '\.docx?$' { return Get-WordText -Path $Path }
            '\.xlsx?$' { return Get-ExcelText -Path $Path }
            '\.pptx?$' { return Get-PowerPointText -Path $Path }
            default    { return "" }
        }
    } catch {
        Log-UnreadableFile -Path $Path -Reason $_.Exception.Message
        return ""
    }
}

function Get-ContentHash {
    param([string]$Content)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Content)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha256.ComputeHash($bytes)
    return -join ($hash | ForEach-Object { "{0:x2}" -f $_ })
}

function Load-EmbeddingCache {
    if (Test-Path $global:OpenAI_RAG_CachePath) {
        Get-Content $global:OpenAI_RAG_CachePath | ConvertFrom-Json
    } else {
        @{}
    }
}

function Save-EmbeddingCache {
    param($cache)
    $cache | ConvertTo-Json -Depth 5 | Set-Content $global:OpenAI_RAG_CachePath
}

function Get-Embedding {
    param([string]$Text)
    $body = @{ input = $Text; model = $global:OpenAI_RAG_EmbeddingModel } | ConvertTo-Json -Depth 10
    $headers = @{ "Authorization" = "Bearer $env:OPENAI_API_KEY"; "Content-Type" = "application/json" }
    $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/embeddings" -Headers $headers -Method Post -Body $body
    return $response.data[0].embedding
}

function Get-FileChunks {
    param([string]$Directory)
    $files = Get-ChildItem -Path $Directory -Include *.txt,*.pdf,*.doc,*.docx,*.xls,*.xlsx,*.ppt,*.pptx -Recurse -ErrorAction SilentlyContinue
    $chunks = @()

    foreach ($file in $files) {
        $content = Get-TextFromFile -Path $file.FullName
        if (-not $content.Trim()) {
            Log-UnreadableFile -Path $file.FullName -Reason "Empty or unreadable content"
            continue
        }

        $paragraphs = $content -split "(`r?`n){2,}"
        foreach ($para in $paragraphs) {
            $clean = $para.Trim()
            if ($clean.Length -gt 50) {
                $chunks += [PSCustomObject]@{
                    file = $file.FullName
                    content = $clean
                    hash = Get-ContentHash -Content $clean
                }
            }
        }
    }
    return $chunks
}

function Get-RelevantChunks {
    param ([string]$Directory, [string]$Query, [int]$TopN = 3)
    $chunks = Get-FileChunks -Directory $Directory
    $cache = Load-EmbeddingCache

    foreach ($chunk in $chunks) {
        if (-not $cache.ContainsKey($chunk.hash)) {
            $embedding = Get-Embedding -Text $chunk.content
            $cache[$chunk.hash] = @{
                file = $chunk.file
                content = $chunk.content
                vector = $embedding
            }
        }
    }

    Save-EmbeddingCache -cache $cache

    $queryVec = Get-Embedding -Text $Query
    $scored = $cache.GetEnumerator() | ForEach-Object {
        $vec = $_.Value.vector
        $score = 0
        for ($i = 0; $i -lt $vec.Count; $i++) {
            $score += $vec[$i] * $queryVec[$i]
        }
        [PSCustomObject]@{
            file = $_.Value.file
            content = $_.Value.content
            score = [math]::Round($score, 4)
        }
    } | Sort-Object -Property score -Descending | Select-Object -First $TopN

    return $scored
}