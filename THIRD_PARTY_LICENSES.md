# Third-Party Open Source Licenses

This project is licensed under Apache License 2.0. The following third-party components are separate works and remain under their own licenses. Commercial use is generally permitted, subject to each license's conditions.

| Component | Pinned version | Purpose | License | Redistribution / notice notes |
|---|---:|---|---|---|
| ReportLab | 5.0.1 | PDF generation | BSD | Commercial use and redistribution permitted; retain applicable copyright/license notice. |
| tkinterdnd2 | 0.6.3 | Tk drag-and-drop wrapper | MIT | Commercial use and redistribution permitted; retain copyright and MIT permission notice. |
| TkDND | bundled by tkinterdnd2 | Native/Tcl drag-and-drop extension | TkDND license terms | Permission is granted for use, copy, modification, distribution, and licensing provided existing copyright notices are retained and the TkDND notice is included verbatim in distributions. |
| watchdog | 6.0.0 | File-system monitoring | Apache-2.0 | Commercial use and redistribution permitted; preserve license/notices required by Apache-2.0. |
| pystray | 0.19.5 | Windows task-tray icon | LGPL-3.0 | Commercial use is permitted. Binary redistribution must comply with LGPLv3, including license/notice obligations and users' rights concerning the LGPL-covered component. Do not relabel pystray as Apache-2.0. |
| Pillow | 12.3.0 | Image/icon handling | MIT-CMU (current Pillow metadata) | Commercial use and redistribution permitted; retain applicable license/copyright notice. |
| PyInstaller | 6.22.3 | Build tool only | GPL-2.0 with PyInstaller exception (plus Apache-2.0 for certain files) | The PyInstaller exception permits distributing generated executable bundles under the application's license, provided bundled dependency licenses are respected. PyInstaller itself is not relicensed as Apache-2.0. |
| Python standard library / Tkinter | Python distribution | Email parsing, GUI, filesystem, Windows integration | Python Software Foundation License; Tcl/Tk terms also apply to bundled runtime components | If distributing a frozen EXE, retain third-party notices supplied by the Python/Tcl/Tk runtime as required. |

## Important build note

`requirements.txt` pins the direct packages used for Community Edition Ver.1.0.0. Python packages may have transitive dependencies. Before every binary release, generate a fresh dependency-license report as described in `docs/BUILD.md` and review any newly introduced package.

## Source references used for the Ver.1.0.0 audit

- ReportLab project metadata: https://pypi.org/project/reportlab/
- tkinterdnd2 project metadata: https://pypi.org/project/tkinterdnd2/
- TkDND license terms: https://github.com/petasis/tkdnd/blob/master/license.terms
- watchdog project metadata: https://pypi.org/project/watchdog/
- pystray project metadata: https://pypi.org/project/pystray/
- Pillow project metadata: https://pypi.org/project/pillow/
- PyInstaller license documentation: https://pyinstaller.org/en/stable/license.html

This file is an engineering inventory, not legal advice. If you distribute binaries commercially or under organization-specific compliance rules, review the exact license files contained in the package versions used to build that release.
