# CHANGELOG

## [2.4.2] - 2026-07-13

### Changed
- **Dynamic Embed Button**: The main embed button now explicitly states what action it will perform (e.g. "Add New Section After..." vs "Embed in Section...") and dynamically changes to blue when "Auto Section" mode is enabled to make the destructive action more obvious.
- **Update Checker Fix**: Fixed a bug where the update checker would incorrectly flag the current version as out of date.


## [2.4.1] - 2026-07-13

### Added
- **Auto-Update Checker**: Added the ability for the app to check GitHub for a newer version.
- **Startup Check Toggle**: Added a "Check updates on startup" option (with a 3-second delay to keep startup fast).
- **Manual Check Button**: Added a "Check for Updates" button at the bottom of the app.
- **Release Notes Viewer**: A new dialog displays the full release notes for a newly available version and provides a download button.

## [2.4.0] - 2026-07-10

### Added
- **1st, 2nd & Last Toggle**: Added a new toggle button to automatically select the first two verses and the last verse of a song.
- **All Toggle**: Added a third, default "All" toggle to explicitly select all verses.
- **Toggle Exclusivity**: The "All", "1st & Last", and "1st, 2nd & Last" toggles are now mutually exclusive to prevent conflicting auto-scanning behaviors. Selecting one will toggle the others off. Manually unchecking the active specific selection will revert to the "All" default.
- **Keyboard Controls**: Added keyboard navigation to speed up embedding. Pressing the Down arrow key from the search box transfers focus to the song list, and pressing Enter on the song list directly embeds the selected song.

### Changed
- **Button Renaming**: Renamed the master slideshow scan button to "Show Verses" and the individual song scan button to "Scan For Verses" to clarify their distinct functions.
- **UI Naming**: Renamed "First & Last" to "1st & Last" to match the style of the other toggles.

## [2.2.0] - 2026-07-06

### Added
- **First & Last Persistent Toggle**: Added a permanent, checkable toggle button for "First & Last" verses that saves its checked state across sessions.
- **Auto-Scan on Embed**: When "First & Last" is enabled, embedding an unscanned song will automatically scan it and select only the first and last verses in a single action.
- **Combined Verse Selection Logic**: Refactored the verse selection detection so both standard embed and auto-append modes fully support the "First & Last" auto-scanning feature.

### Changed
- **Top-Row Navigation**: Relocated the "First & Last" toggle button to the top row next to the "Scan Verses" button for better visibility.
- **Clean Section Naming**: Modified auto-created sections in Auto-Append mode to be named "Song {N}" instead of "Song #{N}" (e.g. "Song 4" instead of "Song #4").
- **Removed All Button**: Removed the redundant "All" button as turning the "First & Last" toggle off acts as the default behavior.

## [2.1.6] - 2026-07-06

### Changed
- **Selected Section Retention**: Embedding a song in normal mode now keeps the currently selected section highlighted rather than advancing automatically to the next section. Auto-append mode still advances to the newly created section at the end of the list.

## [2.1.4] - 2026-07-06

### Added
- **Adjustable Main Divider**: Added a draggable `QSplitter` divide so users can dynamically resize the width of both left and right columns.
- **Resizable Section Columns**: Replaced the custom section list with a native `QTreeWidget`, making the `SECTION`, `TITLE / STATUS`, and `VERSES` columns manually resizable.
- **Persistent UI Sizing State**: The application now saves and restores the custom widths of the main divider and table columns across application launches.
- **Optimized Window Resizing**: Configured the splitter layout to absorb window width adjustments exclusively in the right column (Song Library) so the left column remains at a fixed width unless manually adjusted.

### Changed
- **Compact Verse Selection**: Reduced the padding and removed the minimum width for the verse buttons (`Verse #`, `First & Last`, `All`) to save horizontal space.

### Fixed
- **Selected Text Contrast**: Fixed a styling issue where the "VERSES" text on selected section items remained dark grey and unreadable. It now changes to bright white when highlighted.

## [2.1.0] - 2026-07-06

### Added
- **Auto Append Section Mode**: Added a new "Auto append section" checkbox. When active, embedding a song creates a new sequentially-named "Song #N" section immediately after the currently selected section, rather than replacing or appending to existing slides. This allows users to build a presentation dynamically on the fly during a live service.

## [2.0.1] - 2026-07-06

### Changed
- **Blank Slide Placement**: Reversed blank slide insertion so they are now embedded at the end of a section instead of at the beginning.
- **Verse List Formatting**: Removed the "V" prefix from verse numbers in the Master Presentation section list (displaying e.g. "1, 2" instead of "V1, V2").

## [2.0.0] - 2026-07-02

### Changed
- **UI Architecture & Performance**: Drastically improved master presentation load performance. The app no longer scans for embedded verses immediately upon loading a master presentation.
- **On-Demand Verse Scanning**: Added a dedicated "Scan Verses" button next to the refresh controls, allowing users to scan for embedded verses only when needed.
- **Section Layout Polish**: Fixed jagged column alignment in the master presentation section list. The 'Section' and 'Verses' columns now use strict fixed widths, ensuring the 'Title' column aligns perfectly vertically across all rows to form a neat, professional grid.

## [1.2.2] - 2026-07-02

### Added
- **Master Presentation Verse Tracking**: Added a new 'VERSES' column to the left-side section list. The app now parses the slides currently residing inside the master presentation's "Song" sections to dynamically detect and display exactly which verses (e.g. "V1, V2, V3") are already embedded in the show.


## [1.2.1] - 2026-07-02

### Changed
- **UI Polish**: Consolidated the verse selector panel so the 'First & Last' and 'All' shortcut buttons sit on the exact same row as the dynamically generated verse buttons, saving vertical space.
- **Detailed Scan Results**: The verse information label now dynamically updates as verses are toggled, and displays total slides and chorus detection status (e.g. "12 Slides - 3 Verses - Chorus Detected" or "8 Slides selected").

### Fixed
- Fixed an issue where slides were incorrectly grouped if their title text shape was not the first shape created, or if the title lacked a valid prefix (e.g. lyrics-only slides). The parser now correctly prioritizes the official PowerPoint Title placeholder and automatically inherits the preceding slide's tag when no explicit prefix is found.


## [1.2.0] - 2026-07-02

### Added
- **Verse Selection**: Users can now scan a song PPT to detect its verse/chorus structure and selectively embed only the verses they need.
  - **Scan Verses** button parses slide titles to identify verses (numeric prefix) and choruses (`C` prefix).
  - Dynamically generated **Verse toggle buttons** allow selecting/deselecting individual verses.
  - **First & Last** shortcut button quickly selects only the first and last verses (visible when 3+ verses detected).
  - **All** button re-selects all verses.
  - Choruses are automatically included with their preceding verse.
  - When no verse structure is detected, embedding works as before (all slides).

## [1.1.0] - 2026-07-02

### Added
- Added version number display (`v1.1.0`) to the window title and the bottom right corner of the status bar.
- Added Drag & Drop support: Users can now drag any PowerPoint file (`.ppt`/`.pptx`) from Windows Explorer and drop it anywhere onto the application window to inject it directly into the selected section.
- Added native **Keep Source Formatting** support for one-off file selection and drag-and-drop: Utilizing PowerPoint's native COM copy/paste commands to copy slides exactly as they appear in the source presentation (including custom layouts, color schemes, background designs, and fonts) instead of modifying their styling to match the destination master.
- Restored the music note (`🎵`) icon prefix to section labels, displaying it dynamically only for sections containing "song" (case-insensitive) in their name to distinguish them from other sections (e.g., Sermon, Announcements).

### Changed
- **UI Layout Redesign**: Pivoted the application orientation to a side-by-side (2-column) layout for improved workflow efficiency.
  - **Left Column**: Dedicated to the Master PowerPoint selector and the expanded Sections panel.
  - **Right Column**: Contains the Song Library search, options row, and embedding actions.
- Relocated and renamed the "Inject One-Off PPT" action to the top folder row next to the song library selector as **Select single file**.
- Updated all settings checkboxes to remember their last selected state using `QSettings`.
- Set all settings checkboxes to be checked by default on the first startup.
- Increased the default window width to comfortably accommodate the new dual-column layout.
- Relocated the configuration checkboxes (Replace existing, Insert Blank slide, Require Confirmation, Keep on top) to a secondary row below the action buttons.
- Updated the song/segment title extraction algorithm (`_get_section_title`) to read from the **second slide** in a section, skipping the initial blank placeholder.
- Removed individual row "Select" and "Empty" buttons to declutter the sections list, routing "Empty" actions to the main toolbar button.

### Fixed
- Fixed an issue where the target section dropdown and list selection would reset to the first section after performing a one-off slide injection. Rebuilding the list now correctly preserves the selected slot index.
- Fixed native copying failure by avoiding opening the source file as `ReadOnly`, which blocked COM clipboard operations.

