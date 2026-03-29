# outlook-nav

`outlook-nav` is a small Python wrapper around the Microsoft Outlook COM API. It gives you a cleaner interface for working with mailboxes, folders, and email messages from Python, and it also exposes a Typer-based CLI for common inspection and draft-creation tasks.

This package is intended for environments where the Outlook desktop app is available. The connection layer uses `pywin32`, so it is effectively Windows-only.

## What the module provides

- `OutlookApp`: connect to Outlook, resolve mailboxes, create messages, and work with folders. Supports context manager usage for automatic resource cleanup.
- `Account`: inspect mailbox/account metadata and walk mailbox folder trees.
- `Folder`: list messages (with optional limit and unread-only filters), browse and walk the subfolder tree, create or delete subfolders, and move or delete items.
- `MailItem`: read and set message fields (including HTML body and unread flag), update multiple fields at once, attach files, save, display, send, move, or delete messages.
- `DefaultFolders`: convenient access to default folders including Inbox, Drafts, Sent Mail, Junk, Deleted Items, Outbox, Calendar, Contacts, Journal, Notes, Conflicts, and Local Failures.
- Exception classes: `OutlookError`, `OutlookConnectionError`, `OutlookValidationError`, `EmailValidationError`, and `PathValidationError` for structured error handling.

## Requirements

- Microsoft Outlook installed and configured
- Python with `pywin32`
- `tabulate` and `typer` if you want to use the message table rendering or CLI

## Python examples

### Connect to Outlook

```python
from outlook import OutlookApp

app = OutlookApp()
print(app.list_mailboxes())
```

If you want to target a specific sending account, pass the mailbox SMTP address:

```python
from outlook import OutlookApp

app = OutlookApp(mailbox_address="me@company.com")
print(app.mailbox_account)
```

You can also use the `connect()` classmethod, or manage the connection with a context manager so COM resources are released automatically:

```python
from outlook import OutlookApp

# classmethod — equivalent to OutlookApp() without a mailbox address
app = OutlookApp.connect()

# context manager — calls app.close() on exit
with OutlookApp(mailbox_address="me@company.com") as app:
    print(app.list_mailboxes())
```

### Read messages from the Inbox

```python
from outlook import OutlookApp

app = OutlookApp()
inbox = app.default_folders.inbox

if inbox:
    for message in inbox.list_messages(limit=10):
        print(message.received_time, message.sender_address, message.subject)

    # show only unread messages
    for message in inbox.list_messages(unread_only=True):
        print(message.subject, message.is_unread)
```

### Find a folder by path inside a mailbox

```python
from outlook import OutlookApp

app = OutlookApp(mailbox_address="me@company.com")
account = app.mailbox_account

if account:
    reports = account.find_folder("Inbox/Reports")
    print(reports)
```

### Create and save a draft

```python
from outlook import OutlookApp

app = OutlookApp(mailbox_address="me@company.com")
message = app.create_email()

if message:
    message.to = ["alice@company.com", "bob@company.com"]
    message.cc = "manager@company.com"
    message.bcc = "archive@company.com"
    message.subject = "Weekly status"
    message.body = "Attached is the latest update."
    message.add_attachments("~/Documents/status.xlsx")
    message.save_as_draft()
```

You can also set an HTML body instead of plain text:

```python
if message:
    message.html_body = "<html><body><p>Hello from Python.</p></body></html>"
    message.save_as_draft()
```

To update several fields at once and save in one call, use `update()`:

```python
if message:
    message.update(subject="Updated subject", body="New body text", is_unread=False)
```

### Display or send a message

```python
from outlook import OutlookApp

app = OutlookApp(mailbox_address="me@company.com")
message = app.create_email()

if message:
    message.to = "alice@company.com"
    message.subject = "Hello"
    message.body = "This was created from Python."

    # Open the compose window
    message.display()

    # Or send immediately
    # message.send()
```

### Move a message to another folder

```python
from outlook import OutlookApp

app = OutlookApp(mailbox_address="me@company.com")
account = app.mailbox_account

if account:
    inbox = account.find_folder("Inbox")
    archive = account.find_folder("Inbox/Archive")

    if inbox and archive and inbox.mail_items:
        moved = inbox.mail_items[0].move_to(archive)
        print(moved)
```

### Inspect a message as a dictionary or table

```python
from outlook import OutlookApp

app = OutlookApp()
inbox = app.default_folders.inbox

if inbox and inbox.mail_items:
    message = inbox.mail_items[0]
    print(message.as_dict())
    print(message.as_table())
```

## CLI examples

Run the package as a module:

```bash
python -m outlook --help
```

Use `--verbose` / `-v` to enable detailed debug logging for any command:

```bash
python -m outlook --verbose mailboxes list
```

List visible mailboxes:

```bash
python -m outlook mailboxes list
```

List folders for a mailbox:

```bash
python -m outlook --mailbox "me@company.com" folders list --recursive --max-depth 2
```

List folders starting from a specific subfolder:

```bash
python -m outlook --mailbox "me@company.com" folders list --root "Inbox/Reports" --recursive --max-depth 3
```

List recent Inbox messages:

```bash
python -m outlook messages list Inbox --limit 10 --include-body-preview
```

List only unread messages:

```bash
python -m outlook messages list Inbox --unread-only
```

Create a draft from the CLI:

```bash
python -m outlook drafts create \
  --to "alice@company.com" \
  --subject "Status update" \
  --body "Draft generated from the outlook CLI." \
  --display
```

Create a draft with CC, BCC, and an attachment, then send immediately:

```bash
python -m outlook drafts create \
  --to "alice@company.com" \
  --cc "manager@company.com" \
  --bcc "archive@company.com" \
  --subject "Status update" \
  --body "Please find the report attached." \
  --attachment ~/Documents/report.xlsx \
  --send
```

## Notes

- Email address inputs are validated before being written to Outlook fields.
- Attachment paths must exist on disk.
- When multiple mailboxes are available, mailbox-specific folder lookups work best when you set `mailbox_address` or pass `--mailbox` in the CLI.
- Use `OutlookApp` as a context manager (`with OutlookApp(...) as app:`) or call `app.close()` explicitly to release COM resources when you are done.
- The package exports exception classes (`OutlookConnectionError`, `EmailValidationError`, `PathValidationError`, and others) that you can import and catch for structured error handling.
