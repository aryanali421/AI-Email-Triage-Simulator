# InboxIQ – Email Intelligence

**Your inbox, under control.**

InboxIQ turns scattered email messages into organized priorities, actionable tasks, and draft replies. It is designed to analyze each email separately so that classifications and suggested replies stay specific to the message.

## Core workflow

1. **Import emails** – Upload email samples or paste individual messages.
2. **Separate messages** – When pasting multiple emails, use the separator `--- EMAIL ---`.
3. **Triage settings** – Choose the analysis settings available in the application.
4. **Review results** – Inspect the priorities, extracted actions, and suggested replies.
5. **Take action yourself** – Edit and approve any suggested reply before sending it.

## Supported input formats

- TXT
- PDF
- DOCX
- CSV

The interface description lists a maximum file size of 200 MB per file. Confirm that the deployed app's upload configuration and hosting limits support this size.

## Design principles

- Analyze each email independently.
- Keep suggested replies specific to the email being reviewed.
- Keep the user in control: review everything before taking action.
- Treat generated priorities and replies as suggestions, not automatic decisions.

## Important notes

- Review extracted tasks and classifications for accuracy.
- Do not upload sensitive or confidential email content unless your deployment and data-handling arrangements permit it.
- Email sending is not described as an automated feature; users should review and send replies through their normal email workflow unless an integration is explicitly implemented.

## Technology

Update this section with the exact frameworks, model, and dependencies used by the implementation.

## Author

Aryan Ali
