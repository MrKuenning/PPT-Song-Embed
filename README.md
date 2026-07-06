# Song Embed

A Python and PyQt6 utility to automate the injection/embedding of lyric presentation slides and other slide content directly into a master PowerPoint slideshow presentation.

## Features

- **Automated Embedding**: Instantly replaces slide contents inside targeted sections of an active presentation with slides from external files (e.g., song files in your library).
- **One-Off Injection & Drag-and-Drop**: Directly inject any arbitrary `.ppt` or `.pptx` file (such as announcements, sermons, or lessons) into a selected section by dragging and dropping the file anywhere onto the application window, or by browsing for it manually.
- **Keep Source Formatting**: Native copy/paste integration preserves layout, colors, and design styling of injected files.
- **Section Management**: 
  - List and view all sections inside your active presentation.
  - Insert new sections, delete sections (along with their slides), or swap/reorder sections using **Move Up** and **Move Down**.
  - Empty slides out of sections directly from the tool.
- **Smart Song Detection**: Automatically reads slide text (skipping the initial blank slide in a section) to show you exactly which song/title is currently inside each section (highlighted by a `🎵` music note on song sections).
- **Flexible Options**:
  - Prepend a solid black slide before injected content automatically using the `Insert Blank slide` option.
  - Toggle confirmation dialogs for quick embedding workflows using `Require Confirmation`.
  - Toggle `Keep on top` to prevent the tool window from slipping behind PowerPoint during use.
  - All options remember your last-used state automatically.

## Prerequisites

- Windows OS
- Microsoft PowerPoint installed
- Python 3.x
- Dependencies:
  - `PyQt6`
  - `pywin32` (for PowerPoint COM automation)

## Installation

1. Clone or copy the project files to your system.
2. Install the required Python packages:
   ```bash
   pip install PyQt6 pywin32
   ```

## Compiling to an Executable

To compile the application into a standalone Windows executable (`.exe`):

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Build the app using:
   ```bash
   pyinstaller --noconsole --onefile --icon=icon.ico --add-data "icon.ico;." song_embed.py
   ```
   The compiled `.exe` will be generated in the `dist/` directory.

## Running the Application

To start the application from source, run:
```bash
python song_embed.py
```

## How to Use

1. **Select Master PPT**: Open your master presentation in PowerPoint. Use the dropdown at the top of the Song Embed tool to select it. (Click the refresh button next to it if it doesn't show up right away).
2. **Set Song Library**: Choose the folder where all of your individual song PowerPoints are stored.
3. **Embed Content**:
   - Click on any section in the **Sections** list to select it.
   - Search for a song in the library search bar.
   - Double-click the song or click **Embed** to push it into PowerPoint.
   - Alternatively, click the **Select single file** button to select a file from elsewhere on your system, or **drag & drop a PPT file anywhere onto the Song Embed window** to inject it directly into the selected section.
4. **Manage Sections**: Use the toolbar buttons (`Add`, `Remove`, `Empty`, `Move Up`, `Move Down`) to edit your master slide structure on the fly.
