---
name: word-viewer
description: Open, view, and edit Word documents (.docx, .doc) in the side panel. Trigger when the user wants to view, open, preview, inspect, or edit a Word document.
---

# Word Document Viewer in Side Panel

Whenever the user asks to view, open, inspect, or display a Word document (.docx or .doc):

1. Surface the clickable sidecar pill to open the Word document in the Auxiliary Side Panel:
   - For a specific file:
     `[Open <filename>.docx in Side Panel](sidecar://word-viewer/panel/?file=<encoded_path>)`
     *(Ensure paths with spaces or parentheses are URL-encoded, e.g. `%20`)*
   - For the general Word Viewer:
     `[Open Word Document Viewer](sidecar://word-viewer/panel/)`

2. Display the link prominently so the user can open it with a single click.
