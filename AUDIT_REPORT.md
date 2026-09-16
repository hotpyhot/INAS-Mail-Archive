# Community Edition Conversion Audit

Audit target: `INAS_Mail_Archive_Ver1_7_2.zip`  
Community Edition target: `INAS Mail Archive Community Edition Ver.1.0.0`

## 1. Current implementation analysis

The supplied internal baseline is a Windows Python desktop application. The primary implementation is a single `main.py` (about 4,700 lines) plus locale JSON files, icons, batch build scripts, PowerShell release helpers, and historical specification/change files.

### Languages and technologies

- Python: application, EML parsing, PDF generation orchestration, file operations, settings, UI logic.
- Batch (`.bat`): environment setup, launch, and PyInstaller builds.
- PowerShell (`.ps1`): release ZIP and SHA-256 helper scripts in the internal baseline only.
- JSON: UI localization and runtime settings.

### Direct Python packages found from imports and build files

- `reportlab`: PDF generation.
- `tkinterdnd2`: drag-and-drop support for the Tkinter UI; package contains TkDND native/Tcl components.
- `watchdog`: import-folder monitoring.
- `pystray`: Windows task-tray operation.
- `Pillow` (`PIL`): icon/image handling for the tray.
- `PyInstaller`: EXE build tool.
- Python standard library: `email`, `tkinter`, `pathlib`, `json`, `csv`, `hashlib`, `logging`, `tempfile`, `threading`, `shutil`, `webbrowser`, `winreg`, and others.

### Windows-specific behavior

- `winreg` writes the current-user `Software\\Microsoft\\Windows\\CurrentVersion\\Run` startup entry when enabled.
- `ctypes` is used for a Windows single-instance mutex and to hide the console window after a frozen EXE starts.
- `os.startfile` opens saved folders.
- Windows-installed fonts such as MS Gothic, Yu Gothic, and Meiryo are used when available for PDF output.
- The established PyInstaller mode is `--onefile --console`; `--windowed` is intentionally not used.

### Mail / attachment / PDF flow

- `.eml` is parsed locally with Python's `email.parser.BytesParser` and related email utilities.
- Mail body text and headers are normalized and can be included in a PDF.
- Attachments are extracted from MIME parts and saved to the selected archive folder when requested.
- PDF is generated directly with ReportLab; no Edge, Office, external PDF printer, or cloud PDF API is required by this baseline.
- The application can detect likely external download URLs in mail content and opens the selected URL using the default browser only after a user action.

### External services/programs

- Power Automate / Outlook / OneDrive / SharePoint can be used as an upstream EML delivery workflow but are not program dependencies of the desktop application.
- The default browser is used only when the user opens a detected external link.
- Python is required for source execution; a PyInstaller-built EXE bundles the runtime required for execution.

## 2. Public-information audit

A text scan of the supplied baseline found no real API key, access token, password, private key, organization server path, or non-example email address. Sample addresses use `example.com`. `C:\\Windows\\Fonts` references are Windows system paths and are not user-specific.

The internal baseline contains historical release notes, test documents, internal build commentary, and operational descriptions that are unnecessary for a public repository. Community Edition does not carry those historical internal documents forward.

Community Edition uses dummy public samples such as `example@example.com`, `example.com`, and `C:\\Example\\...` and ignores runtime EML, PDF, logs, settings, history, caches, build output, and virtual environments via `.gitignore`.

## 3. Dependency / license audit

The application source can be released under Apache License 2.0 while preserving each dependency's own license. The notable dependency is `pystray`, which is LGPLv3 and therefore requires explicit compliance when included in a distributed binary. It is retained for Ver.1.0.0 rather than silently removed; `THIRD_PARTY_LICENSES.md` documents the obligation and the release checklist requires reviewing the exact dependency set used for each EXE build.

PyInstaller's documented license exception permits distributing executables built with PyInstaller under the application's selected license, provided the bundled dependencies' licenses are respected.

See `THIRD_PARTY_LICENSES.md` and `docs/BUILD.md`.

## 4. Functions retained in Community Edition Ver.1.0.0

Retained:

- EML import and local parsing.
- PDF archive creation.
- Attachment extraction/selective save.
- Import-folder monitoring.
- Pending-mail list and duplicate detection.
- Exclusion addresses/domains.
- Sender dictionary and naming templates.
- History and destination suggestions.
- Task-tray residency.
- Optional Windows startup registration.
- Japanese/English/Vietnamese locale files.
- External download-link detection/opening.

Removed from the public tree:

- Internal release history and internal baseline/change specification files.
- Internal security-isolation test documentation and duplicate internal test build scripts.
- Internal release ZIP scripts that were tied to the Ver.1.7.x naming line.

No core mail-archive feature was intentionally removed during the initial Community Edition conversion.

## 5. Public repository structure

The public tree is organized as:

- `src/`: application source and centralized application metadata.
- `assets/`: public icon resources.
- `locales/`: language resources.
- `docs/`: build, release, and Power Automate guidance.
- root compliance files: `LICENSE`, `NOTICE`, `THIRD_PARTY_LICENSES.md`, `PRIVACY.md`, `SECURITY.md`, `DISCLAIMER.md`, `CONTRIBUTING.md`.
- public configuration sample: `settings.example.json`.
- `tools/public_release_scan.py`: pre-release source scan helper.

## Conversion fixes made

- Version line reset from internal Ver.1.7.2 to Community Edition Ver.1.0.0.
- App metadata/developer/license/disclaimer centralized in `src/app_info.py`.
- Runtime data directory separated to `%APPDATA%\\INAS\\MailArchiveCommunity\\`.
- Startup registry value separated as `INAS Mail Archive Community Edition`.
- PyInstaller version resource changed to 1.0.0 and Community Edition naming.
- Existing missing import for `urllib.parse.urlparse` was added so external-link detection can execute correctly.
- Source resource lookup adjusted for the new `src/`, `assets/`, and `locales/` layout.
