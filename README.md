# bora.py – Document Recommender

`bora.py` is a **Tkinter-based desktop application** for managing and recommending documents.
It scans directories, extracts text and highlights from multiple formats, and suggests related files based on content similarity.

## ✨ Features

* 📂 **Directory & File Management**

  * Add/remove directories to cache.
  * Toggle directories as active/inactive.
  * Save frequently used files.
  * Incremental cache updates.

* 🔎 **Smart Search**

  * Search across cached documents (supports phrases and keywords).
  * Special `"ayo"` mode for chronological sorting.

* 📑 **Document Info Panel**

  * Shows metadata (last modified, age, publication date).
  * Extracts and displays highlights from:

    * **PDF** (colored highlights via PyMuPDF)
    * **DOCX** (Word highlights)
    * **HTML** (CSS `background-color` spans)

* 🤝 **Recommendations**

  * Suggests similar documents using TF–IDF & cosine similarity.
  * Amplifiers: boost weight of certain keywords.
  * Silencers: downweight less relevant words.

* 🖥️ **GUI Features**

  * Two-pane interface (directory tree + document/recommendations).
  * Zoom mode for chronological browsing.
  * Configurable font size in info panel.
  * Save/load dashboard state (`dashboard.json`).

* 📊 **Persistence**

  * File cache (`cache.json`) for fast reloads.
  * Dashboard state saved across sessions.

## 🛠️ Installation

Main libraries used:

* `tkinter` – GUI framework (bundled with Python)
* `PyMuPDF (fitz)` – PDF parsing & highlights
* `python-docx` – DOCX text & highlight extraction
* `xlrd` – Excel text extraction
* `beautifulsoup4` – HTML parsing
* `scikit-learn` – TF–IDF & cosine similarity
* `winsound` / `winshell` – Windows-only helpers (optional)

## 🚀 Usage

Run the app with:

```bash
python bora_v21.py
```

### Typical workflow

1. **Add directories** containing `.txt`, `.pdf`, `.docx`, `.xls`, `.html`.
2. The app caches documents automatically.
3. Select a file → see metadata, highlights, and recommendations.
4. Use **Amplifiers** to boost important keywords.
5. Use **Silencers** to ignore unhelpful terms.
6. Save frequently used files and toggle directories on/off.
