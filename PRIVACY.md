# Privacy

## Overview

INAS Mail Archive Community Edition processes EML files locally on the Windows PC on which it runs. A source review of Community Edition Ver.1.0.0 found no code that uploads mail content, attachments, settings, history, or logs to a developer-operated server.

The application can detect URLs contained in a mail body. A URL is opened in the user's default web browser only when the user explicitly chooses to open it; after that, handling is performed by the browser and the destination website under their own terms. Power Automate, Outlook, OneDrive, SharePoint, or another synchronization service used to create or transport EML files is external to this application and has its own privacy behavior.

## Local data

By default, application state is stored under `%APPDATA%\INAS\MailArchiveCommunity\`. This can include:

- `settings.json`: folders, user options, exclusion addresses/domains, registered save locations, and related configuration.
- `processed_history.json`: processing history and saved-file information used by the application.
- `sender_dictionary.json`: sender-address based naming data entered or learned by the user.
- `message_id_registry.json`: identifiers used for duplicate detection.
- `logs\inas_mail_archive.log`: operational and error logging.
- other audit logs such as excluded, duplicate, and error-mail logs; these can include EML filenames, sender addresses, reasons, and paths.

The monitored import folder can contain temporary `.eml` files. Depending on settings and user actions, an EML may be moved into an archive folder, deleted from the import folder, or quarantined into `Error` or `Duplicate` subfolders. Saved output can include PDF files and extracted attachments.

## Email and attachments

EML parsing is performed locally using Python's email library. PDF creation is performed locally using ReportLab. Attachments selected for saving are written to the user-selected archive location. The user should not place sensitive mail in folders synchronized to a cloud service unless that is intended.

## Telemetry

Community Edition Ver.1.0.0 does not implement analytics, advertising identifiers, crash-report upload, or developer telemetry.
