import re
from datetime import datetime

UNSET = object()

DIGITAL_ID_ERROR_CODE = -2146893792

SMTP_ADDRESS_SCHEMA = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")
DIGIT_REGEX = re.compile(r"[0-9]+")

_TZ_TIME = datetime.now().astimezone()
_TZ_INFO = _TZ_TIME.tzinfo
TIMEZONE_OFFSET = _TZ_INFO.utcoffset(_TZ_TIME)
