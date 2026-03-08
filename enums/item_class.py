from enum import IntEnum


class OutlookItemClass(IntEnum):
    """
    Outlook Object Class constants.

    Maps to the VBA ``OlObjectClass`` enumeration.
    Deprecated members are marked with a ``# DEPRECATED`` comment.
    """

    # ── Core application objects ──────────────────────────────────────────────
    APPLICATION = 0  # Application object.
    NAMESPACE = 1  # NameSpace object.
    FOLDER = 2  # Folder object.
    FOLDERS = 15  # Folders collection.

    # ── Stores / accounts ─────────────────────────────────────────────────────
    STORE = 107  # Store object.
    STORES = 108  # Stores collection.
    ACCOUNT = 105  # Account object.
    ACCOUNTS = 106  # Accounts collection.
    ACCOUNT_SELECTOR = 180  # AccountSelector object (Outspace UI area).

    # ── Address book ──────────────────────────────────────────────────────────
    ADDRESS_LIST = 7  # AddressList object.
    ADDRESS_LISTS = 20  # AddressLists collection.
    ADDRESS_ENTRY = 8  # AddressEntry object.
    ADDRESS_ENTRIES = 21  # AddressEntries collection.
    EXCHANGE_USER = 110  # ExchangeUser object.
    EXCHANGE_DISTRIBUTION_LIST = 111  # ExchangeDistributionList object.
    DISTRIBUTION_LIST = 69  # ExchangeDistributionList (legacy reference).
    SELECT_NAMES_DIALOG = 109  # SelectNamesDialog object.

    # ── Mail items ────────────────────────────────────────────────────────────
    MAIL_ITEM = 43  # MailItem object.
    REPORT = 46  # ReportItem object.
    REMOTE = 47  # RemoteItem object.
    POST = 45  # PostItem object.
    SHARING = 104  # SharingItem object.

    # ── Meeting items ─────────────────────────────────────────────────────────
    APPOINTMENT = 26  # AppointmentItem object.
    MEETING_REQUEST = 53  # MeetingItem — meeting request.
    MEETING_CANCELLATION = 54  # MeetingItem — meeting cancellation notice.
    MEETING_RESPONSE_NEGATIVE = 55  # MeetingItem — refusal of a meeting request.
    MEETING_RESPONSE_POSITIVE = 56  # MeetingItem — acceptance of a meeting request.
    MEETING_RESPONSE_TENTATIVE = (
        57  # MeetingItem — tentative acceptance of a meeting request.
    )
    MEETING_FORWARD_NOTIFICATION = 181  # MeetingItem — forward notification.
    RECURRENCE_PATTERN = 28  # RecurrencePattern object.
    EXCEPTIONS = 29  # Exceptions collection.
    EXCEPTION = 30  # Exception object.

    # ── Task items ────────────────────────────────────────────────────────────
    TASK = 48  # TaskItem object.
    TASK_REQUEST = 49  # TaskRequestItem object.
    TASK_REQUEST_UPDATE = 50  # TaskRequestUpdateItem object.
    TASK_REQUEST_ACCEPT = 51  # TaskRequestAcceptItem object.
    TASK_REQUEST_DECLINE = 52  # TaskRequestDeclineItem object.

    # ── Other item types ──────────────────────────────────────────────────────
    CONTACT = 40  # ContactItem object.
    JOURNAL = 42  # JournalItem object.
    NOTE = 44  # NoteItem object.
    DOCUMENT = 41  # DocumentItem object.

    # ── Recipients / attachments ──────────────────────────────────────────────
    RECIPIENT = 4  # Recipient object.
    RECIPIENTS = 17  # Recipients collection.
    ATTACHMENT = 5  # Attachment object.
    ATTACHMENTS = 18  # Attachments collection.
    ATTACHMENT_SELECTION = 169  # AttachmentSelection object.

    # ── Items collections ─────────────────────────────────────────────────────
    ITEMS = 16  # Items collection.
    SIMPLE_ITEMS = 179  # SimpleItems collection.
    SELECTION = 74  # Selection object.

    # ── Conversations ─────────────────────────────────────────────────────────
    CONVERSATION = 178  # Conversation object.
    CONVERSATION_HEADER = 182  # ConversationHeader object.

    # ── Actions ───────────────────────────────────────────────────────────────
    ACTION = 32  # Action object.
    ACTIONS = 33  # Actions collection.

    # ── Rules ─────────────────────────────────────────────────────────────────
    RULES = 114  # Rules collection.
    RULE = 115  # Rule object.
    RULE_ACTIONS = 116  # RuleActions collection.
    RULE_ACTION = 117  # RuleAction object.
    MOVE_OR_COPY_RULE_ACTION = 118  # MoveOrCopyRuleAction object.
    SEND_RULE_ACTION = 119  # SendRuleAction object.
    ASSIGN_TO_CATEGORY_RULE_ACTION = 122  # AssignToCategoryRuleAction object.
    PLAY_SOUND_RULE_ACTION = 123  # PlaySoundRuleAction object.
    MARK_AS_TASK_RULE_ACTION = 124  # MarkAsTaskRuleAction object.
    NEW_ITEM_ALERT_RULE_ACTION = 125  # NewItemAlertRuleAction object.

    # ── Rule conditions ───────────────────────────────────────────────────────
    RULE_CONDITIONS = 126  # RuleConditions collection.
    RULE_CONDITION = 127  # RuleCondition object.
    IMPORTANCE_RULE_CONDITION = 128  # ImportanceRuleCondition object.
    CATEGORY_RULE_CONDITION = 130  # CategoryRuleCondition object.
    FORM_NAME_RULE_CONDITION = 131  # FormNameRuleCondition object.
    FROM_RULE_CONDITION = 132  # ToOrFromRuleCondition object.
    SENDER_IN_ADDRESS_LIST_RULE_CONDITION = (
        133  # SenderInAddressListRuleCondition object.
    )
    TEXT_RULE_CONDITION = 134  # TextRuleCondition object.
    ACCOUNT_RULE_CONDITION = 135  # AccountRuleCondition object.
    ADDRESS_RULE_CONDITION = 170  # AddressRuleCondition object.
    FROM_RSS_FEED_RULE_CONDITION = 173  # FromRssFeedRuleCondition object.

    # ── Views ─────────────────────────────────────────────────────────────────
    VIEWS = 79  # Views collection.
    VIEW = 80  # View object (base).
    TABLE_VIEW = 136  # TableView object.
    ICON_VIEW = 137  # IconView object.
    CARD_VIEW = 138  # CardView object.
    CALENDAR_VIEW = 139  # CalendarView object.
    TIMELINE_VIEW = 140  # TimelineView object.
    BUSINESS_CARD_VIEW = 168  # BusinessCardView object.
    PEOPLE_VIEW = 183  # PeopleView object.
    VIEW_FIELDS = 141  # ViewFields collection.
    VIEW_FIELD = 142  # ViewField object.
    VIEW_FONT = 146  # ViewFont object.
    ORDER_FIELD = 144  # OrderField object.
    ORDER_FIELDS = 145  # OrderFields collection.
    AUTO_FORMAT_RULE = 147  # AutoFormatRule object.
    AUTO_FORMAT_RULES = 148  # AutoFormatRules collection.
    COLUMN_FORMAT = 149  # ColumnFormat object.
    COLUMNS = 150  # Columns collection.
    COLUMN = 154  # Column object.

    # ── Navigation pane ───────────────────────────────────────────────────────
    NAVIGATION_PANE = 155  # NavigationPane object.
    NAVIGATION_MODULES = 156  # NavigationModules collection.
    NAVIGATION_MODULE = 157  # NavigationModule object.
    MAIL_MODULE = 158  # MailModule object.
    CALENDAR_MODULE = 159  # CalendarModule object.
    CONTACTS_MODULE = 160  # ContactsModule object.
    TASKS_MODULE = 161  # TasksModule object.
    JOURNAL_MODULE = 162  # JournalModule object.
    NOTES_MODULE = 163  # NotesModule object.
    SOLUTIONS_MODULE = 177  # SolutionsModule object.
    NAVIGATION_GROUPS = 164  # NavigationGroups collection.
    NAVIGATION_GROUP = 165  # NavigationGroup object.
    NAVIGATION_FOLDERS = 166  # NavigationFolders collection.
    NAVIGATION_FOLDER = 167  # NavigationFolder object.

    # ── Outlook bar (legacy nav UI) ───────────────────────────────────────────
    OUTLOOK_BAR_PANE = 63  # OutlookBarPane object.
    OUTLOOK_BAR_STORAGE = 64  # OutlookBarStorage object.
    OUTLOOK_BAR_GROUPS = 65  # OutlookBarGroups collection.
    OUTLOOK_BAR_GROUP = 66  # OutlookBarGroup object.
    OUTLOOK_BAR_SHORTCUTS = 67  # OutlookBarShortcuts collection.
    OUTLOOK_BAR_SHORTCUT = 68  # OutlookBarShortcut object.

    # ── Explorers / inspectors / panes ───────────────────────────────────────
    EXPLORER = 34  # Explorer (main Outlook window) object.
    EXPLORERS = 60  # Explorers collection.
    INSPECTOR = 35  # Inspector (item window) object.
    INSPECTORS = 61  # Inspectors collection.
    PANES = 62  # Panes collection.

    # ── Forms ─────────────────────────────────────────────────────────────────
    PAGES = 36  # Pages collection (form pages).
    FORM_DESCRIPTION = 37  # FormDescription object.
    FORM_REGION = 129  # FormRegion object.
    PROPERTY_PAGE_SITE = 70  # PropertyPageSite object.
    PROPERTY_PAGES = 71  # PropertyPages collection.

    # ── User properties ───────────────────────────────────────────────────────
    USER_PROPERTIES = 38  # UserProperties collection.
    USER_PROPERTY = 39  # UserProperty object.
    USER_DEFINED_PROPERTY = 171  # UserDefinedProperty schema object.
    USER_DEFINED_PROPERTIES = 172  # UserDefinedProperties schema collection.

    # ── Item / property access ────────────────────────────────────────────────
    ITEM_PROPERTIES = 98  # ItemProperties collection.
    ITEM_PROPERTY = 99  # ItemProperty object.
    PROPERTY_ACCESSOR = 112  # PropertyAccessor object (low-level MAPI props).
    TABLE = 120  # Table object (filtered item data).
    ROW = 121  # Row object (record within a Table).

    # ── Storage / sync ────────────────────────────────────────────────────────
    STORAGE_ITEM = 113  # StorageItem object (hidden folder storage).
    SYNC_OBJECT = 72  # SyncObject object.
    SYNC_OBJECTS = 73  # SyncObjects collection.

    # ── Search / results ──────────────────────────────────────────────────────
    SEARCH = 77  # Search object.
    RESULTS = 78  # Results collection.

    # ── Reminders / conflicts ─────────────────────────────────────────────────
    REMINDERS = 100  # Reminders collection.
    REMINDER = 101  # Reminder object.
    CONFLICT = 102  # Conflict object.
    CONFLICTS = 103  # Conflicts collection.

    # ── Calendar sharing ──────────────────────────────────────────────────────
    CALENDAR_SHARING = 151  # CalendarSharing object.

    # ── Categories ────────────────────────────────────────────────────────────
    CATEGORY = 152  # Category object.
    CATEGORIES = 153  # Categories collection.

    # ── Time zones ────────────────────────────────────────────────────────────
    TIME_ZONE = 174  # TimeZone object.
    TIME_ZONES = 175  # TimeZones collection.

    # ── Deprecated ────────────────────────────────────────────────────────────
    LINK = 75  # DEPRECATED.
    LINKS = 76  # DEPRECATED.
    MOBILE = 176  # DEPRECATED.

    @classmethod
    def get(cls, value) -> str:
        try:
            return cls(value).name.capitalize() + "Type"
        except Exception:
            return "_UnknownType"
