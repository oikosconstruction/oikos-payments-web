# OIKOS Construction Payments

Standalone HTML/CSS/JS MVP for construction-payment entry.

## Current mode
The public demo stores test entries only in the browser's localStorage. It does not write to SharePoint and contains no credentials.

## Production target
Microsoft sign-in -> protected backend -> SharePoint -> weekly Excel -> Outlook -> Teams for urgent payments.

## Security
Do not expose a Power Automate HTTP trigger URL or Microsoft 365 secrets in client-side JavaScript.
