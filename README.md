# OpenAI PowerShell RAG Assistant

This toolkit allows you to interact with OpenAI in PowerShell using local documents as context.

## Features

- **Chat interface**: Terminal or dark-themed GUI
- **File types**: TXT, PDF, Word, Excel, PowerPoint
- **Context-aware responses**: Uses embeddings and similarity search
- **Local caching**: Avoids recomputing embeddings
- **Error logging**: See files that couldn’t be read
- **Dark mode GUI**: Includes system prompt and error viewer
- **Multi-line input and timestamps in shell**

## Requirements

- Windows 10+
- Microsoft Office (for Word/Excel/PPT support)
- OpenAI API key

## Setup

1. Set your API key in your terminal:
   ```powershell
   $env:OPENAI_API_KEY = "sk-..."
   ```

2. Run the assistant:
   - GUI: `.\Start-RAG.ps1`
   - Shell: `.\Start-RAG.ps1 -Shell`

## Logging

Unreadable files are logged to:
```
%APPDATA%\OpenAI-RAG\unreadable-files.log
```

Embeddings are cached in:
```
%APPDATA%\OpenAI-RAG\embeddings.json
```