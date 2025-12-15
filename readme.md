# Browse

**Browse** is a lightweight Windows file-explorer-style app (Win32/C++17) with **built-in video playback via libVLC** and a handful of practical, power-user file operations. It was forked from *MediaExplorer* and intentionally **removes all FFmpeg/combine/edit features**—this is meant to stay small and fast.

## Features

### Browsing
- **Drives view** and **folder view**
- Folder view shows **all files** (not just videos)
- Column sorting (directories always shown first)
- Responsive UI while loading large folders (title-bar spinner)
- Optional command-line start folder:  
  `Browse.exe C:\data\new`

### Video playback (libVLC)
- Select one or more video files and play them as a **playlist**
- Seek bar + title bar shows current time / total time
- **Fullscreen** toggle
- **Video metadata columns** (Resolution / Duration) populated quickly when available, then filled in by a background worker

### Video search (videos only)
- **Recursive search** for video files by keyword (case-insensitive)
- Search can be scoped:
  - If you select folders/files before searching, Browse searches **only inside your selection**
- While in Search view, pressing search again adds another keyword and filters results (**AND** semantics)

### File operations
- **Rename** files/folders (F2 or context menu)
- **Delete** files and folders (folders deleted recursively)
- **Cut/Copy/Paste** for **files *and* directories**
- Uses the **system clipboard** (CF_HDROP) so cut/copy from Browse can be pasted in Explorer and vice‑versa
- Paste shows a **progress dialog** with **Cancel**
- Cross-volume moves are handled as copy + delete

### Network drives
- Context menu:
  - **Map Network Drive…**
  - **Disconnect Network Drive…**
- In Drives view, **persistently mapped but disconnected** drives are shown **in red** (displayed as `X:`)
- Disconnected mapped drives are blocked from navigation and offer a single action:
  - **Fix** (disconnect + reconnect to the same `\\server\share`)
- Optional default credentials for map/fix can be read from `browse.ini`

### Logging (optional)
- If enabled, writes `browse.log` to your configured log folder.

## Keyboard shortcuts

### List view (browsing/search)
- **Enter / Double‑click**: open folder OR open file  
  - video file(s) → play in built-in player  
  - non-video file → open with default app (ShellExecute)
- **Left / Backspace**: up one folder (drive root → Drives view)
- **Ctrl+A**: select all
- **F2**: rename selected item
- **Del**: delete selected items (folders deleted recursively)
- **Ctrl+C / Ctrl+X / Ctrl+V**: copy / cut / paste
- **Ctrl+F**: search (recursive) for video files by keyword
- **Right‑click**: context menu (Open/Play/Rename/Cut/Copy/Paste/Delete + network drive actions)

### Playback
- **Esc**: exit playback
- **Enter**: toggle fullscreen
- **Space**: pause
- **Tab**: resume
- **Left / Right**: seek -/+10 seconds (**Shift+Left/Right**: -/+60 seconds)
- **Ctrl+Left / Ctrl+Right**: previous / next item in playlist
- **Up / Down**: volume +/-5
- **Ctrl+G**: open playlist chooser
- **Ctrl+P**: show video properties (optionally uses `ffprobe` for codec details)
- **Del**: queue delete of the current file (applied when you exit playback)
- **Ctrl+R**: queue rename of the current file (applied when you exit playback)
- **Ctrl+C**: queue “copy current file to…” (applied when you exit playback)

> Note: The playback-time file actions (Delete/Rename/Copy-To) are applied when you **exit playback** so playback stays smooth.

## Configuration (browse.ini)

Place `browse.ini` next to `Browse.exe`. The parser is simple `key=value`, case-insensitive. Lines starting with `;` or `#` are comments. Section headers like `[general]` are ignored (allowed, but not required).

Example:

```ini
; browse.ini (next to Browse.exe)

; Optional: show codec details in Ctrl+P “Video properties”
ffprobeAvailable = 1

; Optional: write logs
loggingEnabled = 1
loggingPath    = C:\browse_logs

; Optional: default credentials for mapping/reconnecting network drives
; (leave blank to let Windows prompt interactively)
username = DOMAIN\user
password = your_password
```

### ffprobe notes
If `ffprobeAvailable=1`, ensure `ffprobe.exe` is available on your PATH (or in the working directory). If disabled or unavailable, Browse still shows basic information but may omit codec details.

## Building (Visual Studio 2022)

- **Windows** + **Visual Studio 2022**
- **C++17**, **x64** recommended
- Links against **libVLC** (`libvlc.lib` / `libvlccore.lib`)

Typical steps:

1. Open your solution/project (for example, `Browse.sln`) in Visual Studio 2022.
2. Select **Release | x64**.
3. Ensure include/lib paths point to your libVLC headers and import libraries (many forks keep these under a `vlclib/` folder).
4. Build.

### Runtime requirements (libVLC)
The executable needs the VLC runtime at run time:
- `libvlc.dll`
- `libvlccore.dll`
- the `plugins/` folder

You can provide these next to `Browse.exe`, or use an installed VLC distribution (depending on how you package your app).

## Notes / behavior details

- Folder view shows **all files**; Search view is **videos only**.
- Resolution/Duration columns may appear a moment later for some files due to background metadata fill.
- “Fix” for a broken mapped drive uses the persisted mapping from `HKCU\Network\<Letter>\RemotePath`.

## License

Choose a license for your repository (MIT is a common choice). If this code is derived from another project, ensure you comply with that project’s license and attribution requirements.
