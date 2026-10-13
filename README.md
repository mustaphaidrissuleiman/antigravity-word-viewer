# Word Document Viewer UI Extension for Antigravity

An interactive visual Microsoft Word (.docx, .doc) document viewer that renders directly in the Antigravity 2.0 side panel (Auxiliary Pane).

![Word Viewer Logo](assets/logo.svg)

## Features

- **Side Panel Integration:** Renders smoothly in the right-hand auxiliary pane alongside your conversations.
- **In-Place Text Editing:** Click the **Edit** button (or press `Cmd+E`) to edit text, headings, and paragraphs directly in the document.
- **Rich Formatting Ribbon:** Bold (`Cmd+B`), Italic (`Cmd+I`), Underline (`Cmd+U`), Strikethrough, Headings (H1–H4), Bullet/Numbered lists, Alignment, and Undo/Redo.
- **Direct Save to Disk (`Cmd+S`):** Serializes your edits back into the `.docx` OpenXML structure and writes directly to disk, preserving page geometry, margins, headers, and footers.
- **High-Fidelity Document Layout (Page Mode):** Accurately renders pages, headers, footers, tables, images, footnotes, and styles using vendored `docx-preview` and `JSZip`.
- **Reading / Flow Mode:** Clean responsive fluid reading view powered by `mammoth.js` for comfortable reading in narrow side pane widths.
- **Document Outline (Table of Contents):** Automatically extracts headings (H1, H2, H3...) with one-click jump navigation.
- **Full-Text In-Document Search:** Fast keyword search with match counter and previous/next navigation.
- **Document Inspector & Statistics:** Displays word count, character count, paragraph count, page count, file size, creation date, and core author metadata.
- **Copy & Export Options:** One-click "Copy Text", "Copy Markdown", and "Print / Save as PDF".
- **File Loading:** Open any local Word document via OS file dialog, drag-and-drop, or quick-select from your active workspace.
- **Workspace Discovery:** Automatically detects and indexes `.docx` and `.doc` files inside open workspace project folders.
- **Ask Agent:** One-click integration with Antigravity to summarize, extract action items, review risks, or analyze tables.
- **Self-Contained & 100% Offline:** Vendored dependencies (`docx-preview`, `JSZip`, `mammoth.js`) with zero external network or CDN calls.

## Installation

### Method 1: Global User Plugin (Recommended)

Clone or copy this repository into your Antigravity plugins directory:

- **macOS / Linux:**
  ```bash
  git clone git@github.com:mustaphaidrissuleiman/antigravity-word-viewer.git ~/.gemini/config/plugins/word-viewer
  ```
- **Windows (PowerShell):**
  ```powershell
  git clone https://github.com/mustaphaidrissuleiman/antigravity-word-viewer.git "$env:USERPROFILE\.gemini\config\plugins\word-viewer"
  ```

### Method 2: Project Workspace Plugin

To share the viewer with team members in a specific Git repository:

```bash
mkdir -p .agents/plugins
cp -r /path/to/word-viewer .agents/plugins/word-viewer
```

## Usage

1. Open Antigravity and go to **Customizations > Installed** in the left sidebar.
2. Ensure **Word Document Viewer** is enabled.
3. In any conversation, click the **⋮** menu in the top-right corner, select **Extensions**, and choose **Word Document Viewer**.
4. Or open it via deep-link: `[Open Word Viewer](sidecar://word-viewer/panel/)` or `[Open Document](sidecar://word-viewer/panel/?file=<path>)`.

## License

MIT
