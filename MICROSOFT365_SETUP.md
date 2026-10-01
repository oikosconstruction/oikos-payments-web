# OIKOS Payments — Microsoft 365 direct integration

The application works without Power Automate. Its Vercel backend uses Microsoft Graph with an Entra application identity.

## Production target

Web app -> Vercel server endpoint -> Microsoft Graph -> SharePoint workbook + attachments.
Urgent requests -> Microsoft Graph -> two accounting email recipients.

## Required Vercel environment variables

- OIKOS_MS_TENANT_ID
- OIKOS_MS_CLIENT_ID
- OIKOS_MS_CLIENT_SECRET
- OIKOS_SP_DRIVE_ID
- OIKOS_SP_WORKBOOK_ITEM_ID
- OIKOS_SP_PARENT_FOLDER_ID
- OIKOS_MAIL_SENDER
- OIKOS_ACCOUNTING_EMAIL_1
- OIKOS_ACCOUNTING_EMAIL_2

Known SharePoint targets for this deployment:

- OIKOS_SP_DRIVE_ID = b!0KZbb1RvUE-J7pF1ZPzZNApf4IzoIsBBp-Oex9Cy6tearS2sVy6lR5E5DGYwY9Y6
- OIKOS_SP_WORKBOOK_ITEM_ID = 01L7A7QPSXE4FYY6GO7VEY7QHJ3CI36XZO
- OIKOS_SP_PARENT_FOLDER_ID = 01L7A7QPWIZ7BT4FXNIJF2UD6QZGEZODLC

## Microsoft permissions

Preferred SharePoint permission model: Sites.Selected (Application) with write access granted only to the finance/legal SharePoint site.

Urgent email requires Mail.Send (Application). Admin consent is required.

## Data-safety sequence

1. Validate the request.
2. Upload any attachment into SharePoint.
3. Save a JSON audit record into SharePoint.
4. Download the current workbook.
5. Add the request to tblCurrentPayments.
6. Replace the same workbook only if the SharePoint eTag still matches; retry on concurrent edits.
7. If urgent, send the accounting email and update notification status in the workbook.

The app intentionally avoids the Microsoft workbook-table Graph row endpoint because application-only permissions are not supported for that API.

Teams notification is left for a later approved Teams webhook or delegated Teams integration.
