# Build Guide

## Environment

- Windows 10/11
- VS Code recommended
- Python 3.11 or 3.12 recommended

## Run from source

```bat
py -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python src\main.py
```

## Build EXE

Run `build_exe.bat`. The project intentionally uses PyInstaller `--onefile --console --noupx`. It does **not** use `--windowed`. When running as a frozen Windows EXE, `hide_console_window_for_frozen_app()` hides the console immediately after startup while retaining the console-subsystem build method used by the INAS project.

PyInstaller's license exception permits distribution of executables built from your application under your chosen license, subject to the licenses of bundled dependencies. See `THIRD_PARTY_LICENSES.md`.

## Security-software considerations

No build configuration can guarantee that endpoint protection products will not flag a newly built executable. To reduce avoidable triggers, this project keeps the build reproducible, disables UPX (`--noupx`), avoids self-modifying code and download-and-execute behavior, and documents dependencies. For public releases, publish SHA-256 hashes and, if available, code-sign release binaries.

## Dependency license report

For each release candidate, create a fresh environment and run:

```bat
python -m pip install -r requirements-dev.txt
pip-licenses --format=markdown --with-authors --with-urls > dependency_licenses_build.md
```

Review the generated report before publishing the EXE because transitive dependencies can change when requirements are modified.
