# outlook-nav

`outlook-nav` is a small Python wrapper around the Microsoft Outlook COM API. It gives you a cleaner interface for working with mailboxes, folders, and email messages from Python, and it also exposes a Typer-based CLI for common inspection and draft-creation tasks.

This package is intended for environments where the Outlook desktop app is available. The connection layer uses `pywin32`, so it is effectively Windows-only.

## What the module provides

- `Outlook`: connect to Outlook, resolve accounts, create messages, and work with folders. Supports context manager usage for automatic resource cleanup.
- `Account`: inspect mailbox/account metadata and walk mailbox folder trees.
- `Folder`: list messages (with optional limit and unread-only filters), browse and walk the subfolder tree, create or delete subfolders, and move or delete items. Supports iteration via `for message in folder`.
- `FolderListing`: a dataclass representing a single entry in a folder tree walk, with `path`, `depth`, and `subfolder_count` fields and an `as_row()` helper.
- `MailItem`: read and set message fields (including HTML body and unread flag), update multiple fields at once, attach files, save, show, send, move, export, or delete messages.
- `AddressEntry`: a model representing an Outlook address book entry, exposing `name`, `email_address`, and `user_type`.
- `Account.default_folder()`: resolve default folders such as Inbox, Drafts, Sent Mail, and Junk for a specific account.
- Exception classes: `OutlookError`, `OutlookConnectionError`, `OutlookValidationError`, `EmailValidationError`, and `PathValidationError` for structured error handling.

## Requirements

- Microsoft Outlook installed and configured
- Python with `pywin32`
- `tabulate` and `typer` if you want to use the message table rendering or CLI

## Python examples

### Connect to Outlook

```python
from outlook import Outlook

app = Outlook()
print(app.accounts)
```

If you want to target a specific sending account, pass the mailbox SMTP address:

```python
from outlook import Outlook

app = Outlook(address="me@company.com")
print(app.account)
```

You can manage the connection with a context manager so COM resources are released automatically:

```python
from outlook import Outlook

# context manager — calls app.close() on exit
with Outlook(address="me@company.com") as app:
    print(app.accounts)
```

### Read messages from the Inbox

```python
from outlook import Outlook
from outlook.enums import FolderEnum

app = Outlook()
inbox = app.account.default_folder(FolderEnum.INBOX) if app.account else None

if inbox:
    for message in inbox.list_messages(limit=10):
        print(message.received_at, message.sender_address, message.subject)

    # show only unread messages
    for message in inbox.list_messages(unread_only=True):
        print(message.subject, message.unread)

    # folders also support direct iteration
    for message in inbox:
        print(message.subject)
```

### Find a folder by path inside a mailbox

```python
from outlook import Outlook

app = Outlook(address="me@company.com")
account = app.account

if account:
    reports = account.find_folder("Inbox/Reports")
    print(reports)
```

### Walk and navigate the folder tree

```python
from outlook import Outlook

app = Outlook(address="me@company.com")
account = app.account

if account:
    inbox = account.find_folder("Inbox")
    if inbox:
        # walk the tree up to 3 levels deep
        for entry in inbox.walk(recursive=True, max_depth=3):
            print(entry.path, entry.depth, entry.subfolder_count)

        # get a direct child subfolder by name
        reports = inbox.get_subfolder("Reports")

        # create a new subfolder
        archive = inbox.create_subfolder("Archive")

        # delete a subfolder by name
        inbox.delete_subfolder("OldFolder")
```

### Create and save a draft

```python
from outlook import Outlook

app = Outlook(address="me@company.com")
message = app.new_email()

message.to = ["alice@company.com", "bob@company.com"]
message.cc = "manager@company.com"
message.bcc = "archive@company.com"
message.subject = "Weekly status"
message.body = "Attached is the latest update."
message.add_attachments("~/Documents/status.xlsx")
message.save()
```

You can also set an HTML body instead of plain text using the `html` property:

```python
message.html = "<html><body><p>Hello from Python.</p></body></html>"
message.save()
```

To update several fields at once and save in one call, use `update()`:

```python
message.update(subject="Updated subject", body="New body text", unread=False)
```

### Display or send a message

```python
from outlook import Outlook

app = Outlook(address="me@company.com")
message = app.new_email()

message.to = "alice@company.com"
message.subject = "Hello"
message.body = "This was created from Python."

# Open the compose window
message.show()

# Or send immediately
# message.send()
```

### Move a message to another folder

```python
from outlook import Outlook

app = Outlook(address="me@company.com")
account = app.account

if account:
    inbox = account.find_folder("Inbox")
    archive = account.find_folder("Inbox/Archive")

    if inbox and archive and inbox.mail_items:
        moved = inbox.mail_items[0].move(archive)
        print(moved)
```

### Export (save) a message to disk

```python
from outlook import Outlook
from outlook.enums import FolderEnum

app = Outlook()
inbox = app.account.default_folder(FolderEnum.INBOX) if app.account else None

if inbox and inbox.mail_items:
    message = inbox.mail_items[0]
    success = message.export("~/exports/message.msg")
    print("Saved" if success else "Failed")
```

### Inspect a message as a dictionary or table

```python
from outlook import Outlook
from outlook.enums import FolderEnum

app = Outlook()
inbox = app.account.default_folder(FolderEnum.INBOX) if app.account else None

if inbox and inbox.mail_items:
    message = inbox.mail_items[0]
    print(message.as_dict())
    print(message.as_table())
```

### Inspect sender and recipients

```python
from outlook import Outlook
from outlook.enums import FolderEnum

app = Outlook()
inbox = app.account.default_folder(FolderEnum.INBOX) if app.account else None

if inbox and inbox.mail_items:
    message = inbox.mail_items[0]

    # AddressEntry for the sender (name, email_address, user_type)
    sender = message.sender_entry
    if sender:
        print(sender.name, sender.email_address)

    # Detailed list of all recipients
    for recipient in message.recipients:
        print(recipient["name"], recipient["address"])
```

### Thread and size properties

```python
if inbox and inbox.mail_items:
    message = inbox.mail_items[0]

    # conversation threading
    print(message.thread_id)
    print(message.thread_index)

    # size in bytes and megabytes
    print(message.size)
    print(message.size_mb)

    # parent folder of this message
    print(message.folder)
```

## CLI examples

Run the package as a module:

```bash
python -m outlook --help
```

Use `--verbose` / `-v` to enable detailed debug logging for any command:

```bash
python -m outlook --verbose accounts list
```

List visible accounts:

```bash
python -m outlook accounts list
```

List folders for an account:

```bash
python -m outlook --account "me@company.com" folders list --recursive --max-depth 2
```

List folders starting from a specific subfolder:

```bash
python -m outlook --account "me@company.com" folders list --root "Inbox/Reports" --recursive --max-depth 3
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
- When multiple accounts are available, account-specific folder lookups work best when you set `address` or pass `--account` in the CLI.
- Use `Outlook` as a context manager (`with Outlook(...) as app:`) or call `app.close()` explicitly to release COM resources when you are done.
- The package exports exception classes (`OutlookConnectionError`, `EmailValidationError`, `PathValidationError`, and others) that you can import and catch for structured error handling.
