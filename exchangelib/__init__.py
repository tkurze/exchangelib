from .account import Account, Identity
from .attachments import FileAttachment, ItemAttachment
from .autodiscover import discover
from .configuration import Configuration
from .credentials import DELEGATE, IMPERSONATION, Credentials, OAuth2AuthorizationCodeCredentials, OAuth2Credentials
from .ewsdatetime import UTC, UTC_NOW, EWSDate, EWSDateTime, EWSTimeZone
from .extended_properties import ExtendedProperty
from .folders import DEEP, SHALLOW, Folder, FolderCollection, RootOfHierarchy
from .items import (
    AcceptItem,
    CalendarItem,
    CancelCalendarItem,
    Contact,
    DeclineItem,
    DistributionList,
    ForwardItem,
    Message,
    PostItem,
    ReplyAllToItem,
    ReplyToItem,
    Task,
    TentativelyAcceptItem,
)
from .properties import UID, Attendee, Body, DLMailbox, HTMLBody, ItemId, Mailbox, Room, RoomList
from .protocol import BaseProtocol, FailFast, FaultTolerance, NoVerifyHTTPAdapter, TLSClientAuth
from .restriction import Q
from .settings import OofSettings
from .transport import BASIC, CBA, DIGEST, GSSAPI, NTLM, OAUTH2, SSPI
from .version import Build, Version

__version__ = "4.7.3"

__all__ = [
    "__version__",
    "Account",
    "Identity",
    "FileAttachment",
    "ItemAttachment",
    "discover",
    "Configuration",
    "DELEGATE",
    "IMPERSONATION",
    "Credentials",
    "OAuth2AuthorizationCodeCredentials",
    "OAuth2Credentials",
    "EWSDate",
    "EWSDateTime",
    "EWSTimeZone",
    "UTC",
    "UTC_NOW",
    "ExtendedProperty",
    "Folder",
    "RootOfHierarchy",
    "FolderCollection",
    "SHALLOW",
    "DEEP",
    "AcceptItem",
    "TentativelyAcceptItem",
    "DeclineItem",
    "CalendarItem",
    "CancelCalendarItem",
    "Contact",
    "DistributionList",
    "Message",
    "PostItem",
    "Task",
    "ForwardItem",
    "ReplyToItem",
    "ReplyAllToItem",
    "ItemId",
    "Mailbox",
    "DLMailbox",
    "Attendee",
    "Room",
    "RoomList",
    "Body",
    "HTMLBody",
    "UID",
    "FailFast",
    "FaultTolerance",
    "BaseProtocol",
    "NoVerifyHTTPAdapter",
    "TLSClientAuth",
    "OofSettings",
    "Q",
    "BASIC",
    "DIGEST",
    "NTLM",
    "GSSAPI",
    "SSPI",
    "OAUTH2",
    "CBA",
    "Build",
    "Version",
    "close_connections",
]

# Set a default user agent, e.g. "exchangelib/3.1.1 (python-requests/2.22.0)"
import requests.utils

BaseProtocol.USERAGENT = f"{__name__}/{__version__} ({requests.utils.default_user_agent()})"


def close_connections():
    from .autodiscover import close_connections as close_autodiscover_connections
    from .protocol import close_connections as close_protocol_connections

    close_autodiscover_connections()
    close_protocol_connections()
