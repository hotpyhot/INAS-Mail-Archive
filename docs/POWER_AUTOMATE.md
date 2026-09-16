# Power Automate setup example

This is a general example for delivering EML files to INAS Mail Archive Community Edition. Power Automate is not part of this application. Connector names can vary by Microsoft tenant and language.

## Typical flow

1. Trigger: **Outlook – When a new email arrives (V3)**.
2. Action: **Export email (V2)** to obtain the message content in EML-compatible form.
3. Action: create a file in a synchronized folder such as OneDrive using a unique `.eml` filename.
4. Configure INAS Mail Archive Community Edition's import folder to the corresponding local synchronized folder.
5. Keep the application running with EML monitoring enabled.

## Example paths and addresses

Use dummy or your own values. Do not copy organization-specific values into a public repository.

- Example mail address: `example@example.com`
- Example import path: `C:\Example\MailImport`
- Example archive path: `C:\Example\MailArchive`

## Operational notes

Cloud synchronization can create more than one file-system event while a file is being written. The application includes stabilization and duplicate checks before processing. Test the flow with non-sensitive sample mail before using it for production data.

If the flow handles personal or confidential information, review your Microsoft 365 / Power Automate / OneDrive or SharePoint policies separately.
