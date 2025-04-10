# OpenAI RAG Assistant for PowerShell

A lightweight, offline-friendly **Retrieval-Augmented Generation (RAG)** assistant that runs entirely in **PowerShell**. It lets you chat with OpenAI's GPT models using context pulled from your local files — including **PDFs, Word, Excel, PowerPoint, TXT**, and more.

Supports both:
- **Terminal (CLI)**: Timestamped, styled multi-line input
- **Dark-themed GUI**: With file reindexing and error log viewer

---

## Features

- **Chat with OpenAI GPT-4o or GPT-3.5**
- **Document-augmented responses** (via local embeddings)
- Reads from `My Documents` by default
- Supports:
  - `.txt`, `.pdf`
  - `.doc`, `.docx` (Word)
  - `.xls`, `.xlsx` (Excel)
  - `.ppt`, `.pptx` (PowerPoint)
- **PDF extraction using Windows 10 tools** (no third-party binaries)
- Embeddings are cached for performance
- Unreadable files are logged automatically
- Full support for **GUI and CLI modes**
- No internet dependency other than OpenAI API access

---

## Installation

### Manual Setup

1. [Download the repo ZIP](#) and extract it somewhere (e.g., `C:\OpenAI-RAG`)
2. Open PowerShell in that folder
3. Set your OpenAI API key:
   ```powershell
   $env:OPENAI_API_KEY = "sk-..."
   ```
4. Launch the GUI:
   ```powershell
   .\Start-RAG.ps1
   ```

Or launch the terminal version:
```powershell
.\Start-RAG.ps1 -Shell
```

---

## Requirements

- **Windows 10 or newer**
- **Microsoft Office installed** (for `.docx`, `.xlsx`, `.pptx` support)
- **OpenAI API key**
- PowerShell 5.1+ (preinstalled on most systems)

---

## Default Behavior

- Loads files from your **`Documents`** folder recursively
- Extracts and chunks readable text
- Embeds and caches those chunks
- On each query, uses vector similarity to retrieve the most relevant snippets

---

## File Storage

- **Embedding Cache**: `%APPDATA%\OpenAI-RAG\embeddings.json`
- **Error Log**: `%APPDATA%\OpenAI-RAG\unreadable-files.log`

---

## Supported File Types

| Extension | Description       | Method           |
|-----------|-------------------|------------------|
| `.txt`    | Plain text        | Native PowerShell |
| `.pdf`    | PDF               | Clipboard-based parsing (Windows 10) |
| `.doc/.docx` | Microsoft Word | COM automation (Word.Application) |
| `.xls/.xlsx` | Microsoft Excel | COM automation (Excel.Application) |
| `.ppt/.pptx` | PowerPoint     | COM automation (PowerPoint.Application) |

---

## Example Usage (GPT-4o Shell)

```plaintext
=== OpenAI RAG Shell ===
Context: C:\Users\You\Documents
Type 'exit' to quit, 'reload' to reindex documents.

[YourName 13:42:10]
What are our internal data retention policies?

[Assistant 13:42:12]
Based on the document "DataRetentionPolicy2024.docx", your organization retains sensitive data for 5 years unless legal obligations require longer storage. After this period, secure deletion is mandated using DoD 5220.22-M.

---

[YourName 13:45:31]
Can you summarize the contents of the slide deck from last quarter?

[Assistant 13:45:34]
The presentation "Q4_2023_Review.pptx" highlights a 12% increase in YoY revenue, strong performance in cloud adoption, and notes an upcoming organizational restructuring in Q2 2024.
```

---

## GUI Preview

> *[Screenshot will be included below.]*

---

## Troubleshooting

- **PDFs return blank?** Ensure the Windows clipboard is working properly. Some PDFs may block text selection.
- **Can't read DOCX or XLSX?** Make sure Office is installed and activated.
- **No responses?** Check that `$env:OPENAI_API_KEY` is set and valid.
- **High latency?** GPT-4o is fast, but GPT-4 is slower. You can change models manually in the script.

---

## Roadmap

- [ ] Add model dropdown (GPT-4, GPT-4o, GPT-3.5)
- [ ] Optional context directory selection
- [ ] Enhanced document preview in GUI
- [ ] Offline mode using local LLMs (phi2, Mistral, llama.cpp)
- [ ] MSI/EXE installer with Start Menu shortcut

---

## License

MIT License — fork, modify, and reuse freely.
