# Release Checklist

- [ ] Version in `src/app_info.py`, `version_info.txt`, README, and release filename matches.
- [ ] `python -m compileall src` passes.
- [ ] `requirements.txt` is reviewed and dependency licenses are regenerated.
- [ ] Search tracked text for real email addresses, organization domains, usernames, local paths, API keys, tokens, passwords, private keys, EML content, logs, history, and Power Automate tenant-specific information.
- [ ] `settings.json`, logs, caches, EML, generated PDF, attachments, build folders, and virtual environments are not tracked.
- [ ] Test EML parsing, PDF output, attachment save, duplicate detection, exclusions, tray behavior, startup registration, and import-folder monitoring on Windows.
- [ ] Confirm About displays Community Edition Ver.1.0.0, developer, Apache-2.0, copyright, third-party license notice, and disclaimer.
- [ ] Build with `--onefile --console --noupx`, not `--windowed`.
- [ ] Run endpoint-protection checks in the intended environment.
- [ ] Generate and publish SHA-256 for release ZIP/EXE.
