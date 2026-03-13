# outlook-nav

`outlook-nav` is a small Python wrapper around the Microsoft Outlook COM API. It gives you a cleaner interface for working with mailboxes, folders, and email messages from Python, and it also exposes a Typer-based CLI for common inspection and draft-creation tasks.

This package is intended for environments where the Outlook desktop app is available. The connection layer uses `pywin32`, so it is effectively Windows-only.

## What the module provides

- `OutlookApp`: connect to Outlook, resolve mailboxes, create messages, and work with folders.
- `Account`: inspect mailbox/account metadata and walk mailbox folder trees.
- `Folder`: list messages, browse subfolders, create folders, and move items.
- `MailItem`: read message fields, update drafts, attach files, save, display, send, move, or delete messages.
- `DefaultFolders`: convenient access to common folders like Inbox, Drafts, Sent Mail, and Junk.

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

### Read messages from the Inbox

```python
from outlook import OutlookApp

app = OutlookApp()
inbox = app.default_folders.inbox

if inbox:
    for message in inbox.mail_items[:10]:
        print(message.received_time, message.sender_address, message.subject)
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
    message.subject = "Weekly status"
    message.body = "Attached is the latest update."
    message.add_attachments("~/Documents/status.xlsx")
    message.save()
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

List visible mailboxes:

```bash
python -m outlook mailboxes list
```

List folders for a mailbox:

```bash
python -m outlook --mailbox "me@company.com" folders list --recursive --max-depth 2
```

List recent Inbox messages:

```bash
python -m outlook messages list Inbox --limit 10 --include-body-preview
```

Create a draft from the CLI:

```bash
python -m outlook drafts create \
  --to "alice@company.com" \
  --subject "Status update" \
  --body "Draft generated from the outlook CLI." \
  --display
```

## Notes

- Email address inputs are validated before being written to Outlook fields.
- Attachment paths must exist on disk.
- When multiple mailboxes are available, mailbox-specific folder lookups work best when you set `mailbox_address` or pass `--mailbox` in the CLI.
