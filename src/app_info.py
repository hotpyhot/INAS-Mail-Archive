from __future__ import annotations

APP_NAME = "INAS Mail Archive"
EDITION = "Community Edition"
APP_VERSION = "1.0.0"
BRAND_NAME = "INAS"
BRAND_EXPANDED = "Intelligent Assistant Systems"
TAGLINE = "Making Work Simple."
DEVELOPER = "MASAYUKI HARA"
LICENSE_NAME = "Apache License 2.0"
COPYRIGHT = "Copyright 2026 MASAYUKI HARA"
REPOSITORY_LABEL = "INAS-Mail-Archive"

DISCLAIMER = (
    "This software is provided AS IS, without warranty of operation or fitness for a particular purpose. "
    "To the extent permitted by applicable law, the developer is not liable for damages arising from use "
    "or inability to use the software. Back up important email and attachments before use, and manage email, "
    "attachments, personal information, and confidential information appropriately. This explanatory notice "
    "does not replace the official Apache License 2.0 text in LICENSE."
)

def about_text() -> str:
    return (
        f"{APP_NAME}\n{EDITION}\nVersion {APP_VERSION}\n\n"
        f"{BRAND_NAME}\n{BRAND_EXPANDED}\n{TAGLINE}\n\n"
        f"Developer:\n{DEVELOPER}\n\n"
        f"License:\n{LICENSE_NAME}\n\n"
        f"{COPYRIGHT}\n\n"
        f"GitHub Repository:\n{REPOSITORY_LABEL}\n\n"
        "Third-Party Open Source Licenses:\nSee THIRD_PARTY_LICENSES.md\n\n"
        f"Disclaimer:\n{DISCLAIMER}"
    )
