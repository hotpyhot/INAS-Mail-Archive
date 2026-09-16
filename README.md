# INAS Mail Archive Community Edition

**INAS = Intelligent Assistant Systems**  
**Making Work Simple.**

INAS Mail Archive Community Edition is a Windows desktop application for importing `.eml` mail files, reviewing mail content, creating an archive PDF, and saving selected attachments into organized folders. It can monitor an import folder so EML files created by Outlook/Power Automate workflows can be processed as they arrive.

Community Edition Ver.1.0.0 is a separate public line from the internal `INAS Mail Archive Ver.1.7.x` line. Changes are not automatically synchronized between the two.

## Main features

- Local EML parsing using Python's standard `email` package.
- PDF generation with ReportLab.
- Attachment listing and selective saving.
- Folder monitoring with watchdog.
- Duplicate detection and processing history.
- Sender naming dictionary and exclusion rules.
- Task-tray resident operation.
- Optional Windows startup registration for the current user.
- Japanese, English, and Vietnamese UI resources.
- Detection of likely external file-transfer links in mail text; links open only when the user chooses to open them in the browser.

## Screenshots

Place public screenshots under `docs/images/` before a GitHub release and reference them here. Do not use screenshots containing real mail addresses, message bodies, attachments, organization names, paths, or personal information.

## Supported platform

Windows 10/11. Development is intended primarily for VS Code. Python 3.11 or 3.12 is recommended for source builds.

## Installation for source users

```bat
py -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python src\main.py
```

A template configuration is provided as `settings.example.json`. The application creates the real user configuration under `%APPDATA%\INAS\MailArchiveCommunity\`; do not commit that runtime data.

## Windows EXE build

Run `build_exe.bat`. The standard build uses PyInstaller `--onefile --console --noupx`. `--windowed` is intentionally not used. The frozen application hides its console window immediately after startup. See `docs/BUILD.md` for details and release checks.

## Power Automate integration

A common configuration is:

Outlook → **When a new email arrives (V3)** → **Export email (V2)** → create an `.eml` file in a synchronized folder → INAS Mail Archive monitors that local folder.

General setup guidance is in `docs/POWER_AUTOMATE.md`. Microsoft services and connectors are external services and are not included in this repository.

## Mail data and privacy

Community Edition Ver.1.0.0 processes EML and generates PDF/attachments locally. The reviewed application code does not implement upload of mail data, settings, or logs to a developer-operated server. If the user chooses an external URL detected in a message, the URL is passed to the default browser. Cloud folders or Power Automate flows used by the user have their own data-handling behavior. See `PRIVACY.md`.

Runtime settings, history, sender dictionary, duplicate registry, and logs are stored under `%APPDATA%\INAS\MailArchiveCommunity\`. Archive PDF, attachments, and optionally the original EML are written to locations chosen/configured by the user.

## License

The application source in this repository is licensed under the **Apache License 2.0**. See `LICENSE` and `NOTICE`. Third-party components keep their original licenses; see `THIRD_PARTY_LICENSES.md`.

## Disclaimer

The software is provided AS IS without a guarantee of operation or fitness for a particular purpose. To the extent permitted by applicable law, the developer does not accept liability for damages caused by use or inability to use the software. Keep independent backups of important mail and attachments and handle personal/confidential information appropriately. This explanation does not replace the Apache License 2.0 text. See `DISCLAIMER.md`.

## Security

Do not place API keys, passwords, tokens, private keys, real EML files, personal information, or confidential attachments in public GitHub Issues. See `SECURITY.md` for private-reporting guidance.

## Developer

**MASAYUKI HARA**

INAS — Intelligent Assistant Systems  
Making Work Simple.
