# Song Embed

A lightweight Windows utility to automate the injection/embedding of lyric presentation slides and other slide content directly into a master PowerPoint slideshow presentation. 

Designed specifically to work seamlessly with **The Paperless Hymnal**, this tool allows you to quickly select songs, scan their slide structures, select specific verses, and construct your Sunday service slideshow on the fly.

<img width="1058" height="783" alt="2026-07-06 12_31_35-Song Embed v2 2 0" src="https://github.com/user-attachments/assets/547db27e-2473-4d5f-9acc-7d21adc2087b" />

## Features

- **Quick Song Search Index**: Instantly filter and select songs from your local song ppt library folder using a fast search bar.
- **Song Verse Selection**: Scan songs to detect verse structure and choose specific verses to embed (including a persistent "First & Last" toggle for quick setup).
- **Automated Embedding**: Replace section contents with slides from chosen songs automatically.
- **One-Off PPT Embedding**: Support for injecting one-off slide files (e.g., sermon lessons, announcements) via file browser or direct drag-and-drop.
- **Preserves Source Formatting**: Relies on PowerPoint's native copy-paste flow to preserve exact formatting, layouts, and background designs of the embedded slides.
- **Section Management**: List, add, remove, clear, and reorder presentation sections directly inside the tool.
- **Appending Blank Slides**: Option to automatically insert a solid blank slide at the end of embedded sections to clean up transitions.
- **Flexible Options**: Toggle confirmation dialogs, control window "Keep on top" behavior, and auto-remember your preferences across launches.

## Quick Start (Run the Executable)

If you just want to run the application without installing Python:

1. **Download & Run**: Download the latest compiled `song_embed.exe` from the [Releases page](https://github.com/MrKuenning/PPT-Song-Embed/releases).
2. **Prerequisites**: 
   - Windows OS
   - Microsoft PowerPoint installed and running

## How to Use

> [!IMPORTANT]
> **PowerPoint Sections Requirement**: This application relies heavily on PowerPoint's **Section** feature. Your master presentation *must* contain pre-defined sections for each part of your service (e.g., separate sections for announcements, prayers, individual songs, sermon lessons, etc.) so the tool can target and replace content accurately.

1. **Set Song Library**: Open the song embed program, click on the **Browse** button, and select the folder where all of your individual song PowerPoints (e.g., Paperless Hymnal files) are stored. This scans recursively, so select the root folder that contains all the song files.
<img src="https://github.com/user-attachments/assets/2f0682fe-2611-4ec5-8a96-ff9c4fdbed33" width="49%" alt="Description 2">

2. **Select Master PPT**: Open your master presentation in PowerPoint. Use the dropdown at the top of the Song Embed tool to select it (click the refresh button next to it if it doesn't show up right away). 
   - Once selected, all PowerPoint sections will be listed in the **Sections** table on the left.
   - Any sections containing existing songs will automatically display their song titles alongside a music note `🎵`.
   - Click the optional **Scan Verses** button next to the refresh controls to scan and display which verses are currently embedded in the master presentation.
<img src="https://github.com/user-attachments/assets/245f4785-d342-42af-8a34-b3f83a9ccd793" width="49%" alt="Description 2">

3. **Embed Content**:
   - Click on any section in the **Sections** list to select it.
   - Search for a song in the library search bar.
   - Double-click the song or click **Embed** to push it into PowerPoint.
   - Alternatively, click the **Select single file** button to select a file from elsewhere on your system, or 
> [!TIP]
> Instead of browsing for a single file you can drag & drop a PPT file anywhere onto the Song Embed window to inject it directly into the selected section..

4. **Manage Sections**: Use the toolbar buttons (`Add`, `Remove`, `Empty`, `Move Up`, `Move Down`) to edit your master slide structure on the fly.

## Optional Auto Sections Mode / On-the-fly creation.

For a dynamic song session on-the-fly, you can check the **Auto Section** checkbox. This allows the program to automatically create a new section (e.g. "Song 4") for each song you embed rather than replacing existing slides. 

- This is especially useful for song sessions that do not have a pre-created PowerPoint presentation structure beforehand.
- You can combine this with the **First & Last** toggle to automatically insert the songs with just their first and last verses, or leave it unchecked to insert all verses.

## UI Settings & Buttons Reference

### Left Column (Master PPT Control)

| UI Element / Setting | Action / Purpose |
| :--- | :--- |
| **Select Master PPT** | Dropdown to choose which open PowerPoint presentation to target. |
| **🔄 (Refresh)** | Re-scans active PowerPoint windows and updates the section table. |
| **Scan Verses** | Analyzes existing slides inside master PPT sections to display currently active verses. |
| **Sections Table** | Shows all sections in the active PowerPoint, along with current song title and embedded verses. |
| **➕ Add** | Add a new section to the master PowerPoint after the currently selected section. |
| **❌ Remove** | Delete the selected section and all slides contained within it. |
| **🧹 Empty** | Clear/empty all slides out of the selected section. |
| **⬆️ Move Up / ⬇️ Move Down** | Reorder and swap sections within the master PowerPoint. |

### Right Column (Song Library & Embedding Options)

| UI Element / Setting | Action / Purpose |
| :--- | :--- |
| **📂 Browse** | Select the root folder of your song library. The program will recursively scan the selected folder and sub-folders to build your song library index. |
| **Select single file** | Embeds a one-off PPT file (announcements, lessons) not located in your library folder. |
| **Search Box** | Instantly filters the song library list as you type. |
| **Song List** | Displays the list of songs in the song library. |
| **🔍 Scan Verses** | Analyzes selected song and displays the number of verses in the song. |
| **First & Last** | Toggle to auto-scan and embed only the first and last verses of the selected song. |
| **Target Dropdown** | Dropdown to select which section in the master presentation to embed the song into. |
| **👁 Preview** | Opens the selected song to view its content. |
| **Embed** | Copies the selected song's slides into the active section. |


### Settings

| UI Element / Setting | Action / Purpose |
| :--- | :--- |
| **Keep on top** | Keeps the Song Embed window floating on top of PowerPoint and other apps. |
| **Replace existing** | Checkbox to wipe existing slides in the targeted section before embedding new slides or append slides after the exisitng slides.. |
| **Blank slide** | Checkbox to automatically append a solid blank slide at the end of the embedded content. |
| **Auto Section** | Appends a new section after the selected one, rather than replacing or appending inside it. |
| **Confirmation** | Prompts you with a confirmation warning before performing any slide embed. |


---

## Development & Compiling from Source

If you want to run the application from source code or compile your own binary:

### Prerequisites

- Python 3.x
- Python dependencies:
  ```bash
  pip install PyQt6 pywin32
  ```

### Running from Source

To start the application, run:
```bash
python song_embed.py
```

### Compiling to an Executable

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
