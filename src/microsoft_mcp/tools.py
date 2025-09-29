import base64
import datetime as dt
import logging
import os
import pathlib as pl
from pprint import pformat
import subprocess
from typing import Any
from unittest import result
from urllib.parse import quote
from fastmcp import FastMCP
from . import graph
from .auth import AzureAuthentication
from markitdown import MarkItDown, StreamInfo
from io import BytesIO
from sys import stderr

# Configure logging to both console and file
logger = logging.getLogger(__name__)

# Only configure logging if it hasn't been configured yet
if not logging.getLogger().hasHandlers():
    # Create formatter
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler = logging.StreamHandler(stderr)
    console_handler.setFormatter(formatter)

    # File handler - log to mcp.log in current working directory
    file_handler = logging.FileHandler("mcp.log")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)

    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)

mcp = FastMCP("microsoft-graph-mcp")
# Create a global authentication instance

auth = AzureAuthentication(
    auth_record_file=os.getenv("AZURE_CRED_CACHE_FILE"),
    token_cache_file=os.getenv("AZURE_TOKEN_CACHE_FILE"),
)

# Set the auth instance for the graph module
graph.set_auth_instance(auth)

markitdown = MarkItDown(enable_builtins=True)

FOLDERS = {
    k.casefold(): v
    for k, v in {
        "inbox": "inbox",
        "sent": "sentitems",
        "drafts": "drafts",
        "deleted": "deleteditems",
        "junk": "junkemail",
        "archive": "archive",
    }.items()
}


def convert_to_markdown(html: str, mimetype: str = "text/html") -> str:
    """Convert HTML content to Markdown format."""
    # Use MarkItDown to convert HTML to Markdown
    stream = BytesIO()
    stream.write(html.encode("utf-8"))
    stream.seek(0)
    return markitdown.convert(
        stream, stream_info=StreamInfo(mimetype=mimetype)
    ).text_content


@mcp.tool
def get_user_details(email: str | None = None) -> dict[str, Any]:
    """Get details about a user - either the logged-in user or another user by email address.

    Retrieves user profile information including display name, email, job title, department,
    office location, and other directory information. When no email is provided, returns
    details for the currently signed-in user. When an email is provided, looks up that
    specific user's public profile information.

    Args:
        email: Optional email address of the user to look up. If None, returns current user's details.
               Must be a valid email address format (e.g., "user@company.com").

    Returns:
        User object containing profile information:
        - Basic info: id, displayName, mail, userPrincipalName, givenName, surname
        - Professional: jobTitle, department, companyName, officeLocation, businessPhones
        - Directory info: accountEnabled, userType, createdDateTime
        - When looking up other users, some fields may be limited based on directory permissions

    Examples:
        - get_user_details() - Get current user's profile information
        - get_user_details("colleague@company.com") - Look up specific user's profile
        - get_user_details("manager@company.com") - Get manager's contact information

    Note: Looking up other users requires User.ReadBasic.All permission and the target
    user must be visible in your organization's directory.
    """
    logger.info(f"get_user_details called: email={email}")

    try:
        if email is None:
            # Get current user's details
            result = graph.request("GET", "/me")
            logger.info("get_user_details successful: retrieved current user details")
        else:
            # Look up user by email address
            # Use the /users/{email} endpoint to get user by their email/UPN
            result = graph.request("GET", f"/users/{email}")
            if not result:
                logger.error(
                    f"get_user_details failed: User with email {email} not found"
                )
                raise ValueError(f"User with email {email} not found")
            logger.info(
                f"get_user_details successful: retrieved details for user {email}"
            )

        return result
    except Exception as e:
        logger.error(
            f"get_user_details failed for email={email}: {str(e)}", exc_info=True
        )
        raise


@mcp.prompt
def prepare_work_day():
    return """
    You are a helpful assistant that helps the user to prepare for their work day.
    You can use tools to get information about their calendar, emails, contacts and availability accessing the MS Graph API.
    
    """


@mcp.tool
def is_logged_in() -> bool:
    return auth.exists_valid_token()


@mcp.tool
def login() -> str:
    """Ensure the user is authenticated and return user info.
    Raises an error if authentication fails.

    `login` is required only if tools report errors.
    """

    if not auth.exists_valid_token():
        try:
            auth.get_token()
            return "logged in"
        except Exception as e:
            logger.error(f"login failed: {str(e)}", exc_info=True)
            raise RuntimeError("Login failed, please check authentication settings.")

    else:
        return "already logged in"


@mcp.tool
def list_emails(
    folder: str = "inbox",
    limit: int = 10,
    body_max_length: int = 2000,
    include_body: bool = True,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict[str, Any]]:
    """List emails from a specified folder in the user's mailbox.

    Retrieves emails from common folders like inbox, sent, drafts, etc. Results are ordered by
    received date (most recent first). Use this to get an overview of emails for a specific date range,
    or find recent messages.

    Args:
        folder: Folder name to search in. Options: "inbox", "sent", "drafts", "deleted", "junk", "archive"
        limit: Maximum number of emails to retrieve (1-100, defaults to 10)
        body_max_length: Maximum characters for email body content (default 2000, will truncate if longer)
        include_body: Whether to include email body content (affects response size)
        start_date: Optional start date in ISO format (UTC timezone, e.g., "2024-09-01T00:00:00Z") to filter emails from this date onwards
        end_date: Optional end date in ISO format (UTC timezone, e.g., "2024-09-30T23:59:59Z") to filter emails up to this date

    Returns:
        List of email objects containing id, subject, sender, recipients, date, attachments info,
        and optionally body content. Each email has fields like 'id', 'subject', 'from', 'receivedDateTime'.
        The most recent email (within the specified date range) will be the first included in the results.
        Contains also a deep link to the conversation as `conversation_url` that can be shown to the user to open the email
    Examples:
        - list_emails() - Get 10 most recent inbox emails
        - list_emails(folder="sent", limit=20) - Get 20 recent sent emails
        - list_emails(include_body=False) - Get emails without body content for faster response
        - list_emails(start_date="2024-09-01T00:00:00Z", end_date="2024-09-01T23:59:59Z") - Get emails received on September 1st, 2024
        - list_emails(start_date="2024-08-01T00:00:00Z") - Get emails from August 1st, 2024 onwards
        - list_emails(end_date="2024-08-31T23:59:59Z") - Get emails up to August 31st, 2024
    """
    logger.info(
        f"list_emails called: folder={folder}, limit={limit}, include_body={include_body}, start_date={start_date}, end_date={end_date}"
    )

    try:
        folder_path = FOLDERS.get(folder.casefold(), folder)

        if include_body:
            select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead"
        else:
            select_fields = "id,subject,from,toRecipients,receivedDateTime,hasAttachments,conversationId,isRead"

        params = {
            "$top": min(limit, 100),
            "$select": select_fields,
            "$orderby": "receivedDateTime desc",
        }

        # Add date filtering if provided
        filter_conditions = []
        if start_date:
            filter_conditions.append(f"receivedDateTime ge {start_date}")
        if end_date:
            filter_conditions.append(f"receivedDateTime le {end_date}")

        if filter_conditions:
            params["$filter"] = " and ".join(filter_conditions)

        emails = list(
            graph.request_paginated(
                f"/me/mailFolders/{folder_path}/messages",
                params=params,
                limit=limit,
            )
        )

        for email in emails:
            if include_body:
                # truncate the body
                if "body" in email and "content" in email["body"]:
                    content = email["body"]["content"]
                    if len(content) > body_max_length:
                        email["body"]["content"] = (
                            content[:body_max_length]
                            + f"\n\n[Content truncated - {len(content)} total characters]"
                        )
                        email["body"]["truncated"] = True
                        email["body"]["total_length"] = len(content)
                        logger.info(
                            f"list_emails: body truncated from {len(content)} to {body_max_length} characters"
                        )
            if "conversationId" in email:
                email["conversation_url"] = (
                    f"https://outlook.office.com/mail/deeplink/readconv/{quote(email['conversationId'], safe='')}"
                )

        logger.info(
            f"list_emails successful: retrieved {len(emails)} emails from folder {folder}"
            + (
                f" with date filter start_date={start_date}, end_date={end_date}"
                if start_date or end_date
                else ""
            )
        )
        return emails
    except Exception as e:
        logger.error(f"list_emails failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_email(
    email_id: str,
    include_body: bool = True,
    body_max_length: int = 5000,
    include_attachments: bool = True,
) -> dict[str, Any]:
    """Get detailed information about a specific email by its ID.

    Retrieves complete email details including headers, body content, and attachment metadata.
    Body content can be truncated to manage response size. Use this when you need full email details
    after finding emails with list_emails or search_emails.

    Args:
        email_id: Unique identifier of the email (get from list_emails or search results)
        include_body: Whether to include the email body content in the response
        body_max_length: Maximum characters for body content (default 50000, will truncate if longer)
        include_attachments: Whether to include attachment metadata (names, sizes, types)

    Returns:
        Email object with complete details including:
        - Basic info: id, subject, from, to, cc, receivedDateTime, isRead
        - Body: content as text or markdown, contentType, truncation info if applicable
        - Attachments: list with id, name, size, contentType for each attachment
        - Conversation: conversationId for threading
        - a deep link to the conversation as `conversation_url` that can be shown to the user to open the email

    Examples:
        - get_email("AAMkAD...") - Get full email details
        - get_email(email_id, include_body=False) - Get headers only without body
        - get_email(email_id, body_max_length=1000) - Limit body to 1000 characters
    """
    logger.info(
        f"get_email called: email_id={email_id}, include_body={include_body}, body_max_length={body_max_length}, include_attachments={include_attachments}"
    )

    try:
        params = {}
        if include_attachments:
            params["$expand"] = "attachments($select=id,name,size,contentType)"

        result = graph.request("GET", f"/me/messages/{email_id}", params=params)
        if not result:
            logger.error(f"get_email failed: Email with ID {email_id} not found")
            raise ValueError(f"Email with ID {email_id} not found")

        # Convert HTML to markdown and truncate body if needed
        if include_body and "body" in result and "content" in result["body"]:
            if result["body"]["contentType"].lower() == "html":
                result["body"]["content"] = convert_to_markdown(
                    result["body"]["content"]
                )
                result["body"]["contentType"] = "text/markdown"

            content = result["body"]["content"]
            if len(content) > body_max_length:
                result["body"]["content"] = (
                    content[:body_max_length]
                    + f"\n\n[Content truncated - {len(content)} total characters]"
                )
                result["body"]["truncated"] = True
                result["body"]["total_length"] = len(content)
                logger.info(
                    f"get_email: body truncated from {len(content)} to {body_max_length} characters"
                )
        elif not include_body and "body" in result:
            del result["body"]

        # tidy up to save tokens
        for key in [
            "@odata.context",
            "@odata.etag",
            "parentFolderId",
            "changeKey",
            "internetMessageId",
            "isDeliveryReceiptRequested",
            "isReadReceiptRequested",
        ]:
            if key in result:
                del result[key]
        # add a link to open the whole conversation as "conversation_url"
        if "conversationId" in result:
            result["conversation_url"] = (
                f"https://outlook.office.com/mail/deeplink/readconv/{quote(result['conversationId'], safe='')}"
            )

        # Remove attachment content bytes to reduce size
        if "attachments" in result and result["attachments"]:
            for attachment in result["attachments"]:
                if "contentBytes" in attachment:
                    del attachment["contentBytes"]

        logger.info(f"get_email successful: retrieved email {email_id}")
        return result
    except Exception as e:
        logger.error(
            f"get_email failed for email_id={email_id}: {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def list_events(
    days_ahead: int = 7,
    days_back: int = 0,
    max_body_length: int = 500,
    include_details: bool = False,
) -> list[dict[str, Any]]:
    """List calendar events within a specified date range.

    Retrieves calendar events including recurring event instances. Events are ordered by start time.
    Use this to check upcoming meetings, find events in a date range, or get calendar overview.

    Args:
        days_ahead: Number of days into the future to search (default 7)
        days_back: Number of days into the past to search (default 0 = today onwards)
        include_details: Whether to include full event details like body, attendees, online meeting info

    Returns:
        List of calendar event objects containing:
        - Basic info: id, subject, start/end times, location, organizer (note: All times are in UTC time zone and may require conversion)
        - Details (if include_details=True): body, attendees list, recurrence info, online meeting links
        - Recurring events: individual instances with seriesMasterId for the recurring series

    Examples:
        - list_events() - Get next 7 days of events
        - list_events(days_ahead=30) - Get next 30 days of events
        - list_events(days_back=7, days_ahead=7) - Get events from past week to next week
        - list_events(include_details=False) - Get basic event info only for faster and shorter response
    """
    logger.info(
        f"list_events called: days_ahead={days_ahead}, days_back={days_back}, include_details={include_details}"
    )

    try:
        now = dt.datetime.now(dt.timezone.utc)
        start = (now - dt.timedelta(days=days_back)).isoformat()
        end = (now + dt.timedelta(days=days_ahead)).isoformat()

        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$orderby": "start/dateTime",
            "$top": 100,
        }

        if include_details:
            params["$select"] = (
                "id,subject,start,end,location,body,attendees,organizer,isAllDay,recurrence,onlineMeeting,seriesMasterId"
            )
        else:
            params["$select"] = "id,subject,start,end,location,organizer,seriesMasterId"

        # Use calendarView to get recurring event instances
        events = list(graph.request_paginated("/me/calendarView", params=params))

        # truncate the body content if it exceeds max_body_length
        for event in events:
            if "body" in event:
                if (
                    "content" in event["body"]
                    and len(event["body"]["content"]) > max_body_length
                ):
                    event["body"]["content"] = (
                        event["body"]["content"][:max_body_length] + "..."
                    )

        logger.info(
            f"list_events successful: retrieved {len(events)} events from {start} to {end}"
        )
        return events
    except Exception as e:
        logger.error(f"list_events failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_event(event_id: str) -> dict[str, Any]:
    """Get complete details for a specific calendar event by its ID.

    Retrieves full event information including attendees, recurrence, online meeting details, etc.
    Use this when you need complete event details after finding events with list_events or search_events.

    Args:
        event_id: Unique identifier of the calendar event (get from list_events or search results)

    Returns:
        Complete event object containing:
        - Basic info: id, subject, start/end times, location, isAllDay, organizer
        - Attendees: list with names, email addresses, response status
        - Body: event description/notes
        - Recurrence: pattern info for recurring events
        - Online meeting: Teams/Zoom links and dial-in info if applicable
        - Categories: event categorization tags

    Examples:
        - get_event("AAMkAD...") - Get full details for a specific event
        - Use after list_events() to get complete info about interesting events
    """
    logger.info(f"get_event called: event_id={event_id}")

    try:
        result = graph.request("GET", f"/me/events/{event_id}")
        if not result:
            logger.error(f"get_event failed: Event with ID {event_id} not found")
            raise ValueError(f"Event with ID {event_id} not found")

        logger.info(f"get_event successful: retrieved event {event_id}")
        return result
    except Exception as e:
        logger.error(
            f"get_event failed for event_id={event_id}: {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def check_availability(
    start: str,
    end: str,
    attendees: str | list[str] | None = None,
) -> dict[str, Any]:
    """Check calendar availability for the user and optionally other attendees within a time range.

    Determines free/busy status to help schedule meetings. Shows when people are available,
    busy, or tentatively booked. Useful for finding meeting times that work for everyone.
    All times are in UTC time zone and may require conversion.

    Args:
        start: Start time in ISO format (e.g., "2024-09-02T09:00:00Z" or "2024-09-02T09:00:00")
        end: End time in ISO format
        attendees: Email address(es) of other people to check (optional). Can be single email or list

    Returns:
        Availability information containing:
        - schedules: Array of availability data for each person checked
        - freeBusyViewType: Type of view (e.g., "freeBusy")
        - For each person: email, availability intervals showing free/busy/tentative status
        - availabilityView: Numeric representation of availability (0=free, 1=tentative, 2=busy),
          each number represents a 30-minute interval within the specified time range, starting from the start time.

    Examples:
        - check_availability("2024-09-02T14:00:00Z", "2024-09-02T15:00:00Z") - Check your availability
        - check_availability(start, end, "colleague@company.com") - Check you + one person
        - check_availability(start, end, ["person1@co.com", "person2@co.com"]) - Check multiple people
    """
    logger.info(
        f"check_availability called: start={start}, end={end}, attendees={attendees}"
    )

    try:
        me_info = graph.request("GET", "/me")
        if not me_info or "mail" not in me_info:
            logger.error("check_availability failed: could not get user email address")
            raise ValueError("Failed to get user email address")

        schedules = [me_info["mail"]]
        if attendees:
            attendees_list = [attendees] if isinstance(attendees, str) else attendees
            schedules.extend(attendees_list)
            logger.info(f"check_availability: checking {len(schedules)} schedules")

        payload = {
            "schedules": schedules,
            "startTime": {"dateTime": start, "timeZone": "UTC"},
            "endTime": {"dateTime": end, "timeZone": "UTC"},
            "availabilityViewInterval": 30,
        }

        result = graph.request("POST", "/me/calendar/getSchedule", json=payload)
        if not result:
            logger.error("check_availability failed: no response from server")
            raise ValueError("Failed to check availability")

        logger.info(
            f"check_availability successful: checked availability for {len(schedules)} schedules"
        )
        return result
    except Exception as e:
        logger.error(f"check_availability failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def list_contacts(limit: int = 50) -> list[dict[str, Any]]:
    """List contacts from the user's address book.

    Retrieves personal contacts with names, email addresses, phone numbers, and other details.
    Use this to find contact information, get email addresses for sending messages, or browse contacts.

    Args:
        limit: Maximum number of contacts to retrieve (1-100, defaults to 50)

    Returns:
        List of contact objects containing:
        - Names: givenName, surname, displayName, nickname
        - Email addresses: array of email addresses with labels
        - Phone numbers: businessPhones, homePhones, mobilePhone
        - Addresses: business and home addresses
        - Other: jobTitle, companyName, birthday, notes

    Examples:
        - list_contacts() - Get first 50 contacts
        - list_contacts(limit=100) - Get more contacts
        - Use to find someone's email before sending messages
    """
    logger.info(f"list_contacts called: limit={limit}")

    try:
        params = {"$top": min(limit, 100)}

        contacts = list(
            graph.request_paginated("/me/contacts", params=params, limit=limit)
        )

        logger.info(f"list_contacts successful: retrieved {len(contacts)} contacts")
        return contacts
    except Exception as e:
        logger.error(f"list_contacts failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_contact(contact_id: str) -> dict[str, Any]:
    """Get detailed information for a specific contact by ID.

    Retrieves complete contact details including all phone numbers, email addresses,
    postal addresses, and personal information. Use after finding contacts with list_contacts
    or search_contacts when you need full contact details.

    Args:
        contact_id: Unique identifier of the contact (get from list_contacts or search results)

    Returns:
        Complete contact object containing:
        - Names: givenName, surname, displayName, nickname, title
        - Communications: emailAddresses array, businessPhones, homePhones, mobilePhone
        - Addresses: businessAddress, homeAddress with street, city, state, country, postalCode
        - Professional: jobTitle, companyName, department, officeLocation
        - Personal: birthday, spouseName, children, personalNotes
        - Categories: assigned category tags

    Examples:
        - get_contact("AAMkAD...") - Get full contact details
        - Use to get complete info after finding contact in search results
    """
    logger.info(f"get_contact called: contact_id={contact_id}")

    try:
        result = graph.request("GET", f"/me/contacts/{contact_id}")
        if not result:
            logger.error(f"get_contact failed: Contact with ID {contact_id} not found")
            raise ValueError(f"Contact with ID {contact_id} not found")

        logger.info(f"get_contact successful: retrieved contact {contact_id}")
        return result
    except Exception as e:
        logger.error(
            f"get_contact failed for contact_id={contact_id}: {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def list_files(path: str = "/", limit: int = 50) -> list[dict[str, Any]]:
    """List files and folders in OneDrive at a specified path.

    Browse OneDrive contents to see what files and folders are available. Use this to navigate
    the file system, find documents, or get an overview of stored content.

    Args:
        path: OneDrive path to browse (default "/" for root). Use forward slashes like "Documents/Projects"
        limit: Maximum number of items to retrieve (1-100, defaults to 50)

    Returns:
        List of file/folder objects containing:
        - Basic info: id, name, type (file/folder), size (bytes), modified (timestamp)
        - Download info: download_url for direct file access (for files only)
        - Use 'type' field to distinguish between "file" and "folder"
        - Size is 0 for folders

    Examples:
        - list_files() - List root directory contents
        - list_files(path="Documents") - List contents of Documents folder
        - list_files(path="Pictures/Vacation", limit=100) - Browse specific folder with more results
        - Check 'type' field to see if item is file or folder for navigation
    """
    logger.info(f"list_files called: path={path}, limit={limit}")

    try:
        endpoint = (
            "/me/drive/root/children"
            if path == "/"
            else f"/me/drive/root:/{path}:/children"
        )
        params = {
            "$top": min(limit, 100),
            "$select": "id,name,size,lastModifiedDateTime,folder,file,@microsoft.graph.downloadUrl",
        }

        items = list(graph.request_paginated(endpoint, params=params, limit=limit))

        result = [
            {
                "id": item["id"],
                "name": item["name"],
                "type": "folder" if "folder" in item else "file",
                "size": item.get("size", 0),
                "modified": item.get("lastModifiedDateTime"),
                "download_url": item.get("@microsoft.graph.downloadUrl"),
            }
            for item in items
        ]

        logger.info(
            f"list_files successful: retrieved {len(result)} items from path {path}"
        )
        return result
    except Exception as e:
        logger.error(f"list_files failed for path={path}: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_file(file_id: str, download_path: str) -> dict[str, Any]:
    """Download a file from OneDrive to a local file path.

    Downloads any file from OneDrive to your local computer. Use this after finding files
    with list_files or search_files when you need to access the actual file content.

    Args:
        file_id: Unique identifier of the file (get from list_files or search_files results)
        download_path: Local file path where to save the downloaded file (e.g., "/tmp/document.pdf")

    Returns:
        Download result information:
        - path: Local path where file was saved
        - name: Original filename from OneDrive
        - size_mb: File size in megabytes (rounded to 2 decimals)
        - mime_type: File MIME type (e.g., "application/pdf", "image/jpeg")

    Examples:
        - get_file("AAMkAD...", "/tmp/report.pdf") - Download specific file
        - get_file(file_id, "~/Downloads/document.docx") - Download to Downloads folder
        - Use file_id from list_files() or search_files() results
    """
    logger.info(f"get_file called: file_id={file_id}, download_path={download_path}")

    try:
        import subprocess

        metadata = graph.request("GET", f"/me/drive/items/{file_id}")
        if not metadata:
            logger.error(f"get_file failed: File with ID {file_id} not found")
            raise ValueError(f"File with ID {file_id} not found")

        download_url = metadata.get("@microsoft.graph.downloadUrl")
        if not download_url:
            logger.error(
                f"get_file failed: No download URL available for file {file_id}"
            )
            raise ValueError("No download URL available for this file")

        try:
            subprocess.run(
                ["curl", "-L", "-o", download_path, download_url],
                check=True,
                capture_output=True,
            )

            result = {
                "path": download_path,
                "name": metadata.get("name", "unknown"),
                "size_mb": round(metadata.get("size", 0) / (1024 * 1024), 2),
                "mime_type": (
                    metadata.get("file", {}).get("mimeType") if metadata else None
                ),
            }

            logger.info(
                f"get_file successful: downloaded {result['name']} ({result['size_mb']} MB) to {download_path}"
            )
            return result
        except subprocess.CalledProcessError as e:
            logger.error(f"get_file failed: curl command failed - {e.stderr.decode()}")
            raise RuntimeError(f"Failed to download file: {e.stderr.decode()}")
    except Exception as e:
        logger.error(f"get_file failed for file_id={file_id}: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_attachment(email_id: str, attachment_id: str, save_path: str) -> dict[str, Any]:
    """Download an email attachment to a local file path.

    Downloads attachments from emails (documents, images, etc.) to your local computer.
    Use this after finding emails with attachments via get_email() when you need the actual attachment files.

    Args:
        email_id: Unique identifier of the email containing the attachment
        attachment_id: Unique identifier of the specific attachment (from get_email attachments list)
        save_path: Local file path where to save the attachment (e.g., "/tmp/attachment.pdf")

    Returns:
        Attachment download information:
        - name: Original filename of the attachment
        - content_type: MIME type (e.g., "application/pdf", "image/jpeg", "text/plain")
        - size: File size in bytes
        - saved_to: Absolute local path where file was saved

    Examples:
        - get_attachment(email_id, attachment_id, "/tmp/document.pdf") - Download specific attachment
        - get_attachment(email_id, attachment_id, "~/Downloads/image.jpg") - Save to Downloads
        - First use get_email() to see what attachments are available, then download specific ones
    """
    logger.info(
        f"get_attachment called: email_id={email_id}, attachment_id={attachment_id}, save_path={save_path}"
    )

    try:
        result = graph.request(
            "GET", f"/me/messages/{email_id}/attachments/{attachment_id}"
        )

        if not result:
            logger.error(
                f"get_attachment failed: Attachment {attachment_id} not found in email {email_id}"
            )
            raise ValueError("Attachment not found")

        if "contentBytes" not in result:
            logger.error(
                f"get_attachment failed: Attachment content not available for {attachment_id}"
            )
            raise ValueError("Attachment content not available")

        # Save attachment to file
        path = pl.Path(save_path).expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        content_bytes = base64.b64decode(result["contentBytes"])
        path.write_bytes(content_bytes)

        attachment_result = {
            "name": result.get("name", "unknown"),
            "content_type": result.get("contentType", "application/octet-stream"),
            "size": result.get("size", 0),
            "saved_to": str(path),
        }

        logger.info(
            f"get_attachment successful: saved {attachment_result['name']} ({attachment_result['size']} bytes) to {save_path}"
        )
        return attachment_result
    except Exception as e:
        logger.error(
            f"get_attachment failed for email_id={email_id}, attachment_id={attachment_id}: {str(e)}",
            exc_info=True,
        )
        raise


def _analyze_search_error(error: Exception, request_payload: dict) -> str:
    """Analyze Microsoft Graph Search API errors and provide helpful diagnostics."""
    import httpx

    error_msg = str(error)

    if "400 Bad Request" in error_msg:
        entity_types = request_payload.get("requests", [{}])[0].get("entityTypes", [])
        query_string = (
            request_payload.get("requests", [{}])[0]
            .get("query", {})
            .get("queryString", "")
        )

        suggestions = []

        # Check for problematic entity type combinations
        if len(entity_types) > 1:
            if "event" in entity_types:
                suggestions.append("'event' cannot be combined with other entity types")
            if "person" in entity_types:
                suggestions.append(
                    "'person' cannot be combined with other entity types"
                )
            if "chatMessage" in entity_types and any(
                et in ["driveItem", "site", "list", "listItem"] for et in entity_types
            ):
                suggestions.append(
                    "'chatMessage' cannot be combined with file-related entity types"
                )

        # Check for empty or invalid query
        if not query_string or query_string.isspace():
            suggestions.append(
                "Query string cannot be empty (use '*' for wildcard search)"
            )

        # Check for unsupported entity types
        valid_entities = {
            "message",
            "event",
            "driveItem",
            "site",
            "drive",
            "chatMessage",
            "person",
            "list",
            "listItem",
        }
        invalid_entities = [et for et in entity_types if et not in valid_entities]
        if invalid_entities:
            suggestions.append(f"Invalid entity types: {invalid_entities}")

        if suggestions:
            return f"Bad Request - Possible issues: {'; '.join(suggestions)}"
        else:
            return f"Bad Request - Check entity type combinations and query format"

    elif "401 Unauthorized" in error_msg:
        return "Authentication failed - token may be expired or invalid"
    elif "403 Forbidden" in error_msg:
        return "Insufficient permissions - check that required scopes are granted"
    elif "404 Not Found" in error_msg:
        return "Search endpoint not found - API may not be available in this tenant"
    elif "429 Too Many Requests" in error_msg:
        return "Rate limited - too many requests, please retry later"
    else:
        return f"Unexpected error: {error_msg}"


@mcp.tool
def search_communication(
    query: str,
    limit: int = 10,
    kql_filters: str | None = None,
) -> dict[str, Any]:
    """Search for communication content including emails and Teams messages using Microsoft Search API.

    Searches across Outlook emails and Teams chat/channel messages to find relevant communications.
    Returns metadata and summary information needed to subsequently retrieve full message details
    using get_email, get_chat_message, or get_channel_message tools.

    Args:
        query: Search terms or KQL query string (e.g., "project meeting", "budget discussion")
        limit: Maximum number of results to return (1-100, defaults to 10)
        kql_filters: Additional KQL filters for precise search. Examples:
            - "from:john@company.com" - Emails from specific sender
            - "sent>=2024-01-01" - Messages after specific date
            - "to:manager@company.com" - Emails to specific recipient
            - "IsMentioned:true" - Teams messages where you're mentioned
            - "received>=2024-09-01 AND received<=2024-09-30" - Messages from September 2024
            - "hasAttachments:true" - Messages with attachments

    Returns:
        Search results containing:
        - summary: Search statistics (total results, entity type breakdown)
        - results: Array of matching communication items with:
            - All original fields from Microsoft Graph API (id, subject, from, dates, etc.)
            - entity_type: "message" for emails or "chatMessage" for Teams messages
            - search_rank: Search relevance score
            - search_summary: Brief content preview from search results
            - conversation_url: Deep link for emails (when available)
            - For emails: hitId can be used with get_email(hitId) to retrieve full content
            - For chat messages: chatId and id fields can be used with get_chat_message(chatId, id)

    Examples:
        - search_communication("budget meeting") - Find emails and Teams messages about budget meetings
        - search_communication("project update", kql_filters="from:manager@company.com") - Updates from manager
        - search_communication("quarterly report", kql_filters="sent>=2024-07-01") - Recent quarterly reports
        - search_communication("", kql_filters="IsMentioned:true AND received>=2024-09-01") - Recent mentions

    Note: Use the returned IDs to call get_email(), get_chat_message(), or get_channel_message()
    for complete message content. KQL filters provide precise control over search scope.
    """
    logger.info(
        f"search_communication called: query='{query}', limit={limit}, "
        f"kql_filters='{kql_filters}'"
    )

    try:
        # Communication entity types: emails and Teams messages
        entity_types = ["message", "chatMessage"]

        # Build the search query
        search_query = query.strip()
        if kql_filters:
            if search_query:
                search_query = f"({search_query}) AND ({kql_filters})"
            else:
                search_query = kql_filters

        # Ensure we have a valid search query - Microsoft Graph requires non-empty query
        if not search_query or search_query.isspace():
            search_query = "*"  # Use wildcard for "all content" search

        # Prepare the search request payload
        request_payload = {
            "requests": [
                {
                    "entityTypes": entity_types,
                    "query": {"queryString": search_query},
                    "size": min(limit, 25),  # Microsoft Graph max per request
                    "from": 0,
                }
            ]
        }

        # Note: Microsoft Search API returns a fixed set of fields regardless of what we request
        # For messages, it typically returns: @odata.type, receivedDateTime, subject, bodyPreview, from

        all_results = []
        entity_type_counts = {}
        total_results = 0

        logger.info(
            f"search_communication: Making API request for entity types: {entity_types}"
        )
        logger.info(f"search_communication: Request payload: {request_payload}")

        # Execute search using the graph module's request function
        try:
            result = graph.request("POST", "/search/query", json=request_payload)
            logger.info(f"search_communication: API response received")

            if result and "value" in result:
                for response in result["value"]:
                    if "hitsContainers" in response:
                        for container in response["hitsContainers"]:
                            total_results += container.get("total", 0)

                            if "hits" in container:
                                for hit in container["hits"]:
                                    processed_item = _process_communication_hit(hit)
                                    if processed_item:
                                        all_results.append(processed_item)

                                        # Count entity types
                                        entity_type = processed_item.get(
                                            "entity_type", "unknown"
                                        )
                                        entity_type_counts[entity_type] = (
                                            entity_type_counts.get(entity_type, 0) + 1
                                        )

        except Exception as search_error:
            error_details = _analyze_search_error(search_error, request_payload)
            logger.error(
                f"search_communication API error: {str(search_error)}\nError analysis: {error_details}",
                exc_info=True,
            )
            raise RuntimeError(
                f"Microsoft Graph Search API failed: {error_details}"
            ) from search_error

        # Build response
        response = {
            "summary": {
                "total_results": len(all_results),
                "total_available": total_results,
                "query": query,
                "kql_filters": kql_filters,
                "entity_types_searched": entity_types,
                "entity_type_counts": entity_type_counts,
                "limit_applied": limit,
            },
            "results": all_results[:limit],  # Apply final limit
        }

        logger.info(
            f"search_communication successful: found {len(all_results)} results "
            f"across {len(entity_type_counts)} entity types with query '{search_query}'"
        )
        return response

    except Exception as e:
        logger.error(f"search_communication failed: {str(e)}", exc_info=True)
        error_response = {
            "summary": {
                "total_results": 0,
                "query": query,
                "kql_filters": kql_filters,
                "entity_types_searched": ["message", "chatMessage"],
                "error": str(e),
            },
            "results": [],
        }
        return error_response


def _process_communication_hit(hit: dict[str, Any]) -> dict[str, Any] | None:
    """Process a single communication search hit from Microsoft Graph Search API.

    Extracts all relevant metadata and identifiers needed to subsequently retrieve
    full message content using get_email, get_chat_message, or get_channel_message.
    """
    try:
        resource = hit.get("resource", {})
        hit_id = hit.get("hitId", "")

        if not resource:
            logger.warning(f"No resource found in hit: {hit}")
            return None

        # Make a copy of the resource to avoid modifying the original
        result = dict(resource)

        # Add entity type detection from @odata.type
        odata_type = resource.get("@odata.type", "").lower()
        entity_type = "unknown"
        result["hitId"] = hit_id  # Include hitId
        if "microsoft.graph.message" in odata_type:
            entity_type = "message"
            # For emails, the hitId can be used directly with get_email
            result["id"] = hit_id
        elif "microsoft.graph.chatmessage" in odata_type:
            entity_type = "chatMessage"
            # For chat messages, we need both chatId and id for get_chat_message
            # The id field is already in the resource

        result["entity_type"] = entity_type

        # Add search metadata from the hit
        result["search_rank"] = hit.get("rank", 0)
        result["search_summary"] = hit.get("summary", "")

        # Add conversation URL for emails
        if entity_type == "message" and resource.get("conversationId"):
            result["conversation_url"] = (
                f"https://outlook.office.com/mail/deeplink/readconv/{quote(resource['conversationId'], safe='')}"
            )

        logger.info(
            f"Processed {entity_type} hit: ID={resource.get('id', 'unknown')}, "
            f"Subject/Summary={resource.get('subject', result.get('search_summary', 'No subject'))[:50]}..."
        )

        return result

    except Exception as e:
        logger.warning(f"Failed to process communication hit: {str(e)}")
        return None


def _process_search_hit(
    hit: dict[str, Any],
    include_body: bool,
    body_max_length: int,
) -> dict[str, Any] | None:
    """Process a single search hit from Microsoft Graph Search API.

    Returns the resource data directly from the API with minimal processing,
    just adding entity type detection and body content handling.
    For emails and chat messages, fetches full body content when include_body=True.
    """
    logger.info(f"Processing search hit {pformat(hit)}")
    try:
        resource = hit.get("resource", {})
        if not resource:
            return None

        # Make a copy of the resource to avoid modifying the original
        result = dict(resource)
        logger.info(f"Processing search hit resource: {resource.get('id', 'unknown')}")

        # Add entity type detection from @odata.type
        odata_type = resource.get("@odata.type", "").lower()
        entity_type = "unknown"

        if "message" in odata_type:
            entity_type = "message"
        elif "event" in odata_type:
            entity_type = "event"
        elif "driveitem" in odata_type:
            entity_type = "driveItem"
        elif "chatmessage" in odata_type:
            entity_type = "chatMessage"
        elif "person" in odata_type:
            entity_type = "person"
        elif "site" in odata_type:
            entity_type = "site"
        elif "list" in odata_type and "listitem" not in odata_type:
            entity_type = "list"
        elif "listitem" in odata_type:
            entity_type = "listItem"
        elif "drive" in odata_type:
            entity_type = "drive"

        result["entity_type"] = entity_type

        # Add search metadata from the hit
        result["search_rank"] = hit.get("rank", 0)
        result["search_summary"] = hit.get("summary", "")

        # Add conversation URL for messages
        if entity_type == "message" and resource.get("conversationId"):
            result["conversation_url"] = (
                f"https://outlook.office.com/mail/deeplink/readconv/{quote(resource['conversationId'], safe='')}"
            )

        # For communication items, fetch full body content if requested
        if include_body and entity_type in ["message", "chatMessage"]:
            logger.info(
                f"Fetching full body content for {entity_type} with ID: {resource.get('id', 'unknown')}"
            )
            result = _fetch_full_body_content(result, entity_type, body_max_length)
            logger.info(
                f"After fetch_full_body_content, result has body: {'body' in result}"
            )
        elif not include_body and "body" in result:
            # Remove body if not requested
            logger.info(
                f"Removing body content as include_body=False for {entity_type}"
            )
            del result["body"]

        # Handle content field for files/documents (non-communication items)
        if (
            include_body
            and "content" in result
            and entity_type not in ["message", "chatMessage"]
        ):
            content = result["content"]
            if content and len(content) > body_max_length:
                result["content"] = content[:body_max_length] + "...[truncated]"
                result["content_truncated"] = True
                result["original_content_length"] = len(content)
        elif not include_body and "content" in result:
            del result["content"]

        return result

    except Exception as e:
        logger.warning(f"Failed to process search hit: {str(e)}")
        return None


def _fetch_full_body_content(
    result: dict[str, Any], entity_type: str, body_max_length: int
) -> dict[str, Any]:
    """Process body content for communication items (emails and chat messages).

    The Microsoft Search API does not return full body content or message IDs that can be used
    to fetch full messages. Instead, we use the bodyPreview field that is available in search results.
    """
    try:
        # Microsoft Search API doesn't provide message IDs or full body content
        # Use the bodyPreview field that is available in search results
        body_preview = result.get("bodyPreview", "")

        if body_preview:
            logger.info(f"Using bodyPreview content: {len(body_preview)} characters")

            # Truncate if necessary
            if len(body_preview) > body_max_length:
                content = body_preview[:body_max_length] + "...[truncated]"
                result["body"] = {
                    "content": content,
                    "contentType": "text/plain",
                    "truncated": True,
                    "original_length": len(body_preview),
                    "source": "bodyPreview",
                }
            else:
                result["body"] = {
                    "content": body_preview,
                    "contentType": "text/plain",
                    "source": "bodyPreview",
                }

            logger.info(
                f"Added body content from bodyPreview: {len(result['body']['content'])} characters"
            )
        else:
            logger.warning(f"No bodyPreview available for {entity_type}")
            # Add empty body structure to indicate we tried to fetch body content
            result["body"] = {
                "content": "",
                "contentType": "text/plain",
                "note": "No body preview available from search results",
                "source": "none",
            }

        return result

    except Exception as e:
        logger.warning(f"Failed to process body content: {str(e)}")
        return result


@mcp.tool
def search_files(
    query: str,
    limit: int = 50,
    kql_filters: str | None = None,
) -> dict[str, Any]:
    """Search for files and folders across OneDrive and SharePoint using Microsoft Search API.

    Searches across OneDrive files, SharePoint documents, and sites to find relevant content.
    Supports advanced KQL (Keyword Query Language) filters for precise search results.
    More powerful than browsing folders - searches filenames, document content, metadata, and SharePoint sites.

    Args:
        query: Search terms to find files and sites (e.g., "budget report", "project documents", "quarterly presentation")
        limit: Maximum number of results to return (1-100, defaults to 50)
        kql_filters: Additional KQL filters for precise search. Examples:
            - "filetype:pdf" - Only PDF files
            - "filetype:docx OR filetype:xlsx" - Word or Excel files
            - "author:\"John Smith\"" - Content authored by John Smith
            - "lastModifiedTime>=2024-01-01" - Files modified after specific date
            - "size>=1048576" - Files larger than 1MB (size in bytes)
            - "path:\"/sites/projectsite\"" - Files in specific SharePoint site
            - "contentclass:STS_ListItem_DocumentLibrary" - SharePoint document library items

    Returns:
        Search results containing:
        - summary: Search statistics (total results, entity type breakdown)
        - results: Array of matching file/site items with:
            - All original fields from Microsoft Graph API
            - entity_type: "driveItem" for files/folders or "site" for SharePoint sites
            - search_rank: Search relevance score
            - search_summary: Brief content preview from search
            - For files: id, name, size, lastModifiedDateTime, download_url, webUrl
            - For sites: id, displayName, webUrl, description

    Examples:
        - search_files("presentation") - Find files with "presentation" in name or content
        - search_files("budget 2024", kql_filters="filetype:xlsx") - Find Excel budget files from 2024
        - search_files("", kql_filters="author:\"Sarah Wilson\" AND filetype:pdf") - Sarah's PDF documents
        - search_files("project alpha", kql_filters="lastModifiedTime>=2024-09-01") - Recent project alpha files
        - search_files("quarterly", kql_filters="filetype:pptx OR filetype:pdf") - Quarterly presentations/reports
        - search_files("team site", kql_filters="contentclass:STS_Web") - Find SharePoint team sites

    Note: KQL filters provide precise control over file types, authors, dates, and locations.
    Results include both OneDrive files and SharePoint content for comprehensive coverage.
    """
    logger.info(
        f"search_files called: query='{query}', limit={limit}, kql_filters='{kql_filters}'"
    )

    try:
        # File and site entity types
        entity_types = ["driveItem", "site"]

        # Build the search query
        search_query = query.strip()
        if kql_filters:
            if search_query:
                search_query = f"({search_query}) AND ({kql_filters})"
            else:
                search_query = kql_filters

        # Ensure we have a valid search query - Microsoft Graph requires non-empty query
        if not search_query or search_query.isspace():
            search_query = "*"  # Use wildcard for "all content" search

        # Prepare the search request payload
        request_payload = {
            "requests": [
                {
                    "entityTypes": entity_types,
                    "query": {"queryString": search_query},
                    "size": min(limit, 25),  # Microsoft Graph max per request
                    "from": 0,
                }
            ]
        }

        all_results = []
        entity_type_counts = {}
        total_results = 0

        logger.info(
            f"search_files: Making API request for entity types: {entity_types}"
        )

        # Execute search using the graph module's request function
        try:
            result = graph.request("POST", "/search/query", json=request_payload)
            logger.info(f"search_files: API response received")

            if result and "value" in result:
                for response in result["value"]:
                    if "hitsContainers" in response:
                        for container in response["hitsContainers"]:
                            total_results += container.get("total", 0)

                            if "hits" in container:
                                for hit in container["hits"]:
                                    processed_item = _process_search_hit(
                                        hit,
                                        include_body=False,  # Files don't need body content
                                        body_max_length=0,
                                    )
                                    if processed_item:
                                        all_results.append(processed_item)

                                        # Count entity types
                                        entity_type = processed_item.get(
                                            "entity_type", "unknown"
                                        )
                                        entity_type_counts[entity_type] = (
                                            entity_type_counts.get(entity_type, 0) + 1
                                        )

        except Exception as search_error:
            error_details = _analyze_search_error(search_error, request_payload)
            logger.error(
                f"search_files API error: {str(search_error)}\nError analysis: {error_details}",
                exc_info=True,
            )
            raise RuntimeError(
                f"Microsoft Graph Search API failed: {error_details}"
            ) from search_error

        # Build response
        response = {
            "summary": {
                "total_results": len(all_results),
                "total_available": total_results,
                "query": query,
                "kql_filters": kql_filters,
                "entity_types_searched": entity_types,
                "entity_type_counts": entity_type_counts,
                "limit_applied": limit,
            },
            "results": all_results[:limit],  # Apply final limit
        }

        logger.info(
            f"search_files successful: found {len(all_results)} results "
            f"across {len(entity_type_counts)} entity types with query '{search_query}'"
        )
        return response

    except Exception as e:
        logger.error(f"search_files failed: {str(e)}", exc_info=True)
        error_response = {
            "summary": {
                "total_results": 0,
                "query": query,
                "kql_filters": kql_filters,
                "entity_types_searched": ["driveItem", "site"],
                "error": str(e),
            },
            "results": [],
        }
        return error_response


@mcp.tool
def search_emails(
    query: str,
    limit: int = 50,
    folder: str | None = None,
) -> list[dict[str, Any]]:
    """Search for emails using text queries across subjects, content, and metadata.

    Find emails by searching subject lines, body content, sender/recipient names, and other metadata.
    Can search across all emails or within a specific folder. More powerful than browsing folders.

    Args:
        query: Search terms (e.g., "meeting notes", "project update", sender name, subject keywords)
        limit: Maximum number of results to return (1-100, defaults to 50)
        folder: Optional folder to search within ("inbox", "sent", "drafts", etc.). If None, searches all emails

    Returns:
        List of matching email objects containing:
        - Basic info: id, subject, from, toRecipients, receivedDateTime, isRead
        - Indicators: hasAttachments, conversationId
        - Body content (if available)
        - Results ranked by relevance to search query
        - a deep link to the conversation as `conversation_url` that can be shown to the user to open the email

    Examples:
        - search_emails("project alpha") - Find emails about "project alpha" anywhere
        - search_emails("meeting", folder="inbox") - Find meeting emails only in inbox
        - search_emails("emails received today") - Finds emails from today
        - search_emails("john.doe@company.com") - Find emails from/to specific person
        - search_emails("budget approval") - Find emails about budget approvals
    """
    logger.info(
        f"search_emails called: query='{query}', limit={limit}, folder={folder}"
    )

    try:
        if folder:
            # For folder-specific search, use the traditional endpoint
            folder_path = FOLDERS.get(folder.casefold(), folder)
            endpoint = f"/me/mailFolders/{folder_path}/messages"

            params = {
                "$search": f'"{query}"',
                "$top": min(limit, 100),
                "$select": "id,subject,from,toRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead",
            }

            result = list(graph.request_paginated(endpoint, params=params, limit=limit))
            for email in result:
                if "conversationId" in email:
                    email["conversation_url"] = (
                        f"https://outlook.office.com/mail/deeplink/readconv/{quote(email['conversationId'], safe='')}"
                    )
                # tidy up to save tokens
                for key in [
                    "@odata.context",
                    "@odata.etag",
                    "parentFolderId",
                    "changeKey",
                    "internetMessageId",
                    "isDeliveryReceiptRequested",
                    "isReadReceiptRequested",
                ]:
                    if key in email:
                        del email[key]

            logger.info(
                f"search_emails successful: found {len(result)} emails in folder '{folder}' matching '{query}'"
            )
            return result

        result = list(graph.search_query(query, ["message"], limit))
        for email in result:
            if "conversationId" in email:
                email["conversation_url"] = (
                    f"https://outlook.office.com/mail/deeplink/readconv/{quote(email['conversationId'], safe='')}"
                )
            # tidy up to save tokens
            for key in [
                "@odata.context",
                "@odata.etag",
                "parentFolderId",
                "changeKey",
                "internetMessageId",
                "isDeliveryReceiptRequested",
                "isReadReceiptRequested",
            ]:
                if key in email:
                    del email[key]

        logger.info(
            f"search_emails successful: found {len(result)} emails matching '{query}'"
        )
        return result
    except Exception as e:
        logger.error(
            f"search_emails failed for query='{query}', folder={folder}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def search_events(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for calendar events using text queries across titles, content, and metadata.

    Find calendar events by searching event titles, descriptions, locations, and attendee information.
    Useful for finding specific meetings, events with certain keywords, or events in particular locations.

    Args:
        query: Search terms (e.g., "team meeting", "conference room", "project review", attendee names)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching calendar event objects containing:
        - Basic info: id, subject, start/end times, location, organizer
        - Details: body/description, attendees, isAllDay status
        - Meeting info: onlineMeeting links if applicable
        - Recurrence: seriesMasterId for recurring events
        - Results ranked by relevance to search query

    Examples:
        - search_events("standup") - Find daily standup meetings
        - search_events("conference room A") - Find events in specific room
        - search_events("john smith") - Find events with John Smith as organizer/attendee
        - search_events("quarterly review") - Find quarterly review meetings
    """
    logger.info(f"search_events called: query='{query}', limit={limit}")

    try:
        events = list(graph.search_query(query, ["event"], limit))

        logger.info(
            f"search_events successful: found {len(events)} events matching '{query}'"
        )
        return events
    except Exception as e:
        logger.error(
            f"search_events failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def search_contacts(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for contacts using text queries across names, email addresses, and other fields.

    Find contacts by searching names, email addresses, phone numbers, company names, and other
    contact information. More efficient than browsing all contacts when looking for specific people.

    Args:
        query: Search terms (e.g., person name, email address, company name, phone number)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching contact objects containing:
        - Names: givenName, surname, displayName, nickname
        - Communications: emailAddresses, businessPhones, homePhones, mobilePhone
        - Professional: jobTitle, companyName, department
        - Addresses: business and home address information
        - Results ranked by relevance to search query

    Examples:
        - search_contacts("john") - Find contacts with "john" in their name
        - search_contacts("microsoft") - Find contacts who work at Microsoft
        - search_contacts("john.doe@company.com") - Find contact with specific email
        - search_contacts("555-0123") - Find contact with specific phone number
    """
    logger.info(f"search_contacts called: query='{query}', limit={limit}")

    try:
        params = {
            "$search": f'"{query}"',
            "$top": min(limit, 100),
        }

        contacts = list(
            graph.request_paginated("/me/contacts", params=params, limit=limit)
        )

        logger.info(
            f"search_contacts successful: found {len(contacts)} contacts matching '{query}'"
        )
        return contacts
    except Exception as e:
        logger.error(
            f"search_contacts failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def list_chat_messages(
    limit: int = 10,
    body_max_length: int = 2000,
    include_body: bool = True,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict[str, Any]]:
    """List recent chat messages from all 1:1 and group chats.

    Retrieves messages from Microsoft Teams chats, ordered by most recent first. This includes
    both one-on-one chats and group chat conversations. Use this to get an overview of recent
    chat activity across all conversations.

    Args:
        limit: Maximum number of messages to retrieve (1-100, defaults to 10)
        body_max_length: Maximum characters for message body content (default 2000, will truncate if longer)
        include_body: Whether to include message body content (affects response size)
        start_date: Optional start date in ISO format (UTC timezone, e.g., "2024-09-01T00:00:00Z") to filter messages from this date onwards
        end_date: Optional end date in ISO format (UTC timezone, e.g., "2024-09-30T23:59:59Z") to filter messages up to this date

    Returns:
        List of message objects containing:
        - Basic info: id, chatId, messageType, createdDateTime, from (sender info)
        - Content: body with text/html content and optionally attachments info
        - Chat context: chat title, chat type (oneOnOne/group), participant count
        - Web URL for opening the message in Teams client
        - The most recent message will be first in the results

    Examples:
        - list_chat_messages() - Get 10 most recent chat messages
        - list_chat_messages(limit=50) - Get more recent messages
        - list_chat_messages(include_body=False) - Get messages without body content for faster response
        - list_chat_messages(start_date="2024-09-01T00:00:00Z") - Get messages from September 1st onwards
    """
    logger.info(
        f"list_chat_messages called: limit={limit}, include_body={include_body}, start_date={start_date}, end_date={end_date}"
    )

    try:
        # First get all chats
        chats = list(graph.request_paginated("/me/chats", params={"$top": 50}))

        all_messages = []

        for chat in chats:
            chat_id = chat["id"]

            # Build message query parameters
            if include_body:
                select_fields = "id,messageType,createdDateTime,from,body,attachments"
            else:
                select_fields = "id,messageType,createdDateTime,from"

            params = {
                "$top": min(limit, 100),
                "$select": select_fields,
                "$orderby": "createdDateTime desc",
            }

            # Add date filtering if provided
            filter_conditions = []
            if start_date:
                filter_conditions.append(f"createdDateTime ge {start_date}")
            if end_date:
                filter_conditions.append(f"createdDateTime le {end_date}")

            if filter_conditions:
                params["$filter"] = " and ".join(filter_conditions)

            try:
                messages = list(
                    graph.request_paginated(
                        f"/me/chats/{chat_id}/messages",
                        params=params,
                        limit=min(
                            limit, 20
                        ),  # Limit per chat to avoid too many results
                    )
                )

                for message in messages:
                    # Add chat context to each message
                    message["chatId"] = chat_id
                    message["chatTopic"] = chat.get("topic", "")
                    message["chatType"] = chat.get("chatType", "")
                    message["webUrl"] = chat.get("webUrl", "")

                    # Process message body
                    if (
                        include_body
                        and "body" in message
                        and "content" in message["body"]
                    ):
                        content = message["body"]["content"]
                        if message["body"].get("contentType") == "html":
                            content = convert_to_markdown(content)
                            message["body"]["contentType"] = "text/markdown"
                            message["body"]["content"] = content

                        if len(content) > body_max_length:
                            message["body"]["content"] = (
                                content[:body_max_length]
                                + f"\n\n[Content truncated - {len(content)} total characters]"
                            )
                            message["body"]["truncated"] = True
                            message["body"]["total_length"] = len(content)

                all_messages.extend(messages)

            except Exception as chat_error:
                logger.warning(
                    f"Failed to get messages from chat {chat_id}: {str(chat_error)}"
                )
                continue

        # Sort all messages by creation time (most recent first) and limit results
        all_messages.sort(key=lambda x: x.get("createdDateTime", ""), reverse=True)
        result = all_messages[:limit]

        logger.info(
            f"list_chat_messages successful: retrieved {len(result)} messages from {len(chats)} chats"
        )
        return result
    except Exception as e:
        logger.error(f"list_chat_messages failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def list_channel_messages(
    team_id: str | None = None,
    channel_id: str | None = None,
    limit: int = 10,
    body_max_length: int = 2000,
    include_body: bool = True,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict[str, Any]]:
    """List recent messages from Microsoft Teams channels.

    Retrieves messages from team channels, ordered by most recent first. If no team/channel is specified,
    gets messages from all accessible channels. Use this to monitor channel activity and discussions.

    Args:
        team_id: Optional ID of specific team to get messages from
        channel_id: Optional ID of specific channel (requires team_id)
        limit: Maximum number of messages to retrieve (1-100, defaults to 10)
        body_max_length: Maximum characters for message body content (default 2000, will truncate if longer)
        include_body: Whether to include message body content (affects response size)
        start_date: Optional start date in ISO format (UTC timezone, e.g., "2024-09-01T00:00:00Z") to filter messages from this date onwards
        end_date: Optional end date in ISO format (UTC timezone, e.g., "2024-09-30T23:59:59Z") to filter messages up to this date

    Returns:
        List of message objects containing:
        - Basic info: id, messageType, createdDateTime, from (sender info)
        - Content: body with text/html content and optionally attachments info
        - Channel context: team name, channel name, channel description
        - Web URL for opening the message in Teams client
        - The most recent message will be first in the results

    Examples:
        - list_channel_messages() - Get 10 most recent messages from all channels
        - list_channel_messages(team_id="abc123") - Get messages from specific team's channels
        - list_channel_messages(team_id="abc123", channel_id="def456") - Get messages from specific channel
        - list_channel_messages(limit=50, include_body=False) - Get more messages without body content
    """
    logger.info(
        f"list_channel_messages called: team_id={team_id}, channel_id={channel_id}, limit={limit}, include_body={include_body}"
    )

    try:
        all_messages = []

        if team_id and channel_id:
            # Get messages from specific channel
            teams_channels = [(team_id, channel_id)]
        elif team_id:
            # Get all channels from specific team
            channels = list(graph.request_paginated(f"/teams/{team_id}/channels"))
            teams_channels = [(team_id, ch["id"]) for ch in channels]
        else:
            # Get messages from all teams and channels user has access to
            teams = list(graph.request_paginated("/me/joinedTeams"))
            teams_channels = []
            for team in teams:
                try:
                    channels = list(
                        graph.request_paginated(f"/teams/{team['id']}/channels")
                    )
                    teams_channels.extend([(team["id"], ch["id"]) for ch in channels])
                except Exception as team_error:
                    logger.warning(
                        f"Failed to get channels from team {team['id']}: {str(team_error)}"
                    )
                    continue

        for t_id, c_id in teams_channels:
            # Build message query parameters
            if include_body:
                select_fields = "id,messageType,createdDateTime,from,body,attachments"
            else:
                select_fields = "id,messageType,createdDateTime,from"

            params = {
                "$top": min(limit, 50),
                "$select": select_fields,
                "$orderby": "createdDateTime desc",
            }

            # Add date filtering if provided
            filter_conditions = []
            if start_date:
                filter_conditions.append(f"createdDateTime ge {start_date}")
            if end_date:
                filter_conditions.append(f"createdDateTime le {end_date}")

            if filter_conditions:
                params["$filter"] = " and ".join(filter_conditions)

            try:
                # Get team and channel info for context
                team_info = graph.request("GET", f"/teams/{t_id}")
                channel_info = graph.request("GET", f"/teams/{t_id}/channels/{c_id}")

                messages = list(
                    graph.request_paginated(
                        f"/teams/{t_id}/channels/{c_id}/messages",
                        params=params,
                        limit=min(limit, 20),  # Limit per channel
                    )
                )

                for message in messages:
                    # Add channel context to each message
                    message["teamId"] = t_id
                    message["channelId"] = c_id
                    message["teamName"] = (
                        team_info.get("displayName", "") if team_info else ""
                    )
                    message["channelName"] = (
                        channel_info.get("displayName", "") if channel_info else ""
                    )
                    message["webUrl"] = (
                        channel_info.get("webUrl", "") if channel_info else ""
                    )

                    # Process message body
                    if (
                        include_body
                        and "body" in message
                        and "content" in message["body"]
                    ):
                        content = message["body"]["content"]
                        if message["body"].get("contentType") == "html":
                            content = convert_to_markdown(content)
                            message["body"]["contentType"] = "text/markdown"
                            message["body"]["content"] = content

                        if len(content) > body_max_length:
                            message["body"]["content"] = (
                                content[:body_max_length]
                                + f"\n\n[Content truncated - {len(content)} total characters]"
                            )
                            message["body"]["truncated"] = True
                            message["body"]["total_length"] = len(content)

                all_messages.extend(messages)

            except Exception as channel_error:
                logger.warning(
                    f"Failed to get messages from team {t_id}, channel {c_id}: {str(channel_error)}"
                )
                continue

        # Sort all messages by creation time (most recent first) and limit results
        all_messages.sort(key=lambda x: x.get("createdDateTime", ""), reverse=True)
        result = all_messages[:limit]

        logger.info(
            f"list_channel_messages successful: retrieved {len(result)} messages from {len(teams_channels)} channels"
        )
        return result
    except Exception as e:
        logger.error(f"list_channel_messages failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_chat_message(chat_id: str, message_id: str) -> dict[str, Any]:
    """Get detailed information about a specific chat message by its ID.

    Retrieves complete message details including content, attachments, reactions, and replies.
    Use this when you need full message details after finding messages with list_chat_messages
    or search_chat_messages.

    Args:
        chat_id: Unique identifier of the chat containing the message
        message_id: Unique identifier of the specific message

    Returns:
        Complete message object containing:
        - Basic info: id, messageType, createdDateTime, lastModifiedDateTime, from (sender)
        - Content: body with full text/html content, contentType
        - Attachments: list with attachment details if any
        - Reactions: emoji reactions and who reacted
        - Chat context: chat info, participants
        - Web URL for opening in Teams client

    Examples:
        - get_chat_message(chat_id, message_id) - Get full message details
        - Use after finding interesting messages with list_chat_messages()
    """
    logger.info(f"get_chat_message called: chat_id={chat_id}, message_id={message_id}")

    try:
        # Get the message details
        message = graph.request("GET", f"/me/chats/{chat_id}/messages/{message_id}")
        if not message:
            logger.error(
                f"get_chat_message failed: Message {message_id} not found in chat {chat_id}"
            )
            raise ValueError(f"Message {message_id} not found in chat {chat_id}")

        # Get chat context
        try:
            chat_info = graph.request("GET", f"/me/chats/{chat_id}")
            if chat_info:
                message["chatContext"] = {
                    "topic": chat_info.get("topic", ""),
                    "chatType": chat_info.get("chatType", ""),
                    "webUrl": chat_info.get("webUrl", ""),
                }
        except Exception as context_error:
            logger.warning(f"Could not get chat context: {str(context_error)}")

        # Convert HTML body to markdown if needed
        if "body" in message and "content" in message["body"]:
            if message["body"].get("contentType") == "html":
                message["body"]["content"] = convert_to_markdown(
                    message["body"]["content"]
                )
                message["body"]["contentType"] = "text/markdown"

        # Get message replies if any
        try:
            replies = list(
                graph.request_paginated(
                    f"/me/chats/{chat_id}/messages/{message_id}/replies"
                )
            )
            if replies:
                message["replies"] = replies
                message["replyCount"] = len(replies)
        except Exception as replies_error:
            logger.warning(f"Could not get message replies: {str(replies_error)}")

        logger.info(f"get_chat_message successful: retrieved message {message_id}")
        return message
    except Exception as e:
        logger.error(
            f"get_chat_message failed for chat_id={chat_id}, message_id={message_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def get_channel_message(
    team_id: str, channel_id: str, message_id: str
) -> dict[str, Any]:
    """Get detailed information about a specific channel message by its ID.

    Retrieves complete message details including content, attachments, reactions, and replies.
    Use this when you need full message details after finding messages with list_channel_messages
    or search_channel_messages.

    Args:
        team_id: Unique identifier of the team containing the channel
        channel_id: Unique identifier of the channel containing the message
        message_id: Unique identifier of the specific message

    Returns:
        Complete message object containing:
        - Basic info: id, messageType, createdDateTime, lastModifiedDateTime, from (sender)
        - Content: body with full text/html content, contentType
        - Attachments: list with attachment details if any
        - Reactions: emoji reactions and who reacted
        - Channel context: team name, channel name, web URL
        - Web URL for opening in Teams client

    Examples:
        - get_channel_message(team_id, channel_id, message_id) - Get full message details
        - Use after finding interesting messages with list_channel_messages()
    """
    logger.info(
        f"get_channel_message called: team_id={team_id}, channel_id={channel_id}, message_id={message_id}"
    )

    try:
        # Get the message details
        message = graph.request(
            "GET", f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}"
        )
        if not message:
            logger.error(f"get_channel_message failed: Message {message_id} not found")
            raise ValueError(f"Message {message_id} not found in channel {channel_id}")

        # Get team and channel context
        try:
            team_info = graph.request("GET", f"/teams/{team_id}")
            channel_info = graph.request(
                "GET", f"/teams/{team_id}/channels/{channel_id}"
            )

            if team_info and channel_info:
                message["channelContext"] = {
                    "teamName": team_info.get("displayName", ""),
                    "channelName": channel_info.get("displayName", ""),
                    "channelDescription": channel_info.get("description", ""),
                    "webUrl": channel_info.get("webUrl", ""),
                }
        except Exception as context_error:
            logger.warning(f"Could not get channel context: {str(context_error)}")

        # Convert HTML body to markdown if needed
        if "body" in message and "content" in message["body"]:
            if message["body"].get("contentType") == "html":
                message["body"]["content"] = convert_to_markdown(
                    message["body"]["content"]
                )
                message["body"]["contentType"] = "text/markdown"

        # Get message replies if any
        try:
            replies = list(
                graph.request_paginated(
                    f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"
                )
            )
            if replies:
                message["replies"] = replies
                message["replyCount"] = len(replies)
        except Exception as replies_error:
            logger.warning(f"Could not get message replies: {str(replies_error)}")

        logger.info(f"get_channel_message successful: retrieved message {message_id}")
        return message
    except Exception as e:
        logger.error(
            f"get_channel_message failed for team_id={team_id}, channel_id={channel_id}, message_id={message_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def search_chat_messages(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for chat messages using text queries across message content and metadata.

    Find chat messages by searching message content, sender names, and other metadata across
    all accessible 1:1 and group chats. More powerful than browsing recent messages when looking
    for specific content or conversations.

    Args:
        query: Search terms (e.g., "project update", "meeting notes", sender name, keywords)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching message objects containing:
        - Basic info: id, chatId, messageType, createdDateTime, from (sender)
        - Content: body with text content matching the search query
        - Chat context: chat topic, chat type, participant info
        - Web URL for opening in Teams client
        - Results ranked by relevance to search query

    Examples:
        - search_chat_messages("project alpha") - Find chat messages about "project alpha"
        - search_chat_messages("john smith") - Find messages from/mentioning John Smith
        - search_chat_messages("meeting tomorrow") - Find messages about upcoming meetings
        - search_chat_messages("budget approval") - Find chat messages about budget approvals
    """
    logger.info(f"search_chat_messages called: query='{query}', limit={limit}")

    try:
        # Use Microsoft Graph search API to search across chat messages
        result = list(graph.search_query(query, ["chatMessage"], limit))

        # Process results to add chat context
        for message in result:
            # Extract chat ID from message if available
            if "chatId" in message:
                chat_id = message["chatId"]
                try:
                    chat_info = graph.request("GET", f"/me/chats/{chat_id}")
                    if chat_info:
                        message["chatTopic"] = chat_info.get("topic", "")
                        message["chatType"] = chat_info.get("chatType", "")
                        message["webUrl"] = chat_info.get("webUrl", "")
                except Exception as context_error:
                    logger.warning(
                        f"Could not get chat context for chat {chat_id}: {str(context_error)}"
                    )

            # Convert HTML to markdown if needed
            if "body" in message and "content" in message["body"]:
                if message["body"].get("contentType") == "html":
                    message["body"]["content"] = convert_to_markdown(
                        message["body"]["content"]
                    )
                    message["body"]["contentType"] = "text/markdown"

        logger.info(
            f"search_chat_messages successful: found {len(result)} messages matching '{query}'"
        )
        return result
    except Exception as e:
        logger.error(
            f"search_chat_messages failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def search_channel_messages(
    query: str,
    team_id: str | None = None,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for channel messages using text queries across message content and metadata.

    Find channel messages by searching message content, sender names, and other metadata across
    all accessible team channels or within a specific team. More powerful than browsing recent
    messages when looking for specific content or discussions.

    Args:
        query: Search terms (e.g., "project update", "announcement", sender name, keywords)
        team_id: Optional ID of specific team to search within (searches all teams if not provided)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching message objects containing:
        - Basic info: id, teamId, channelId, messageType, createdDateTime, from (sender)
        - Content: body with text content matching the search query
        - Channel context: team name, channel name, web URL
        - Results ranked by relevance to search query

    Examples:
        - search_channel_messages("project alpha") - Find channel messages about "project alpha"
        - search_channel_messages("announcement", team_id="abc123") - Find announcements in specific team
        - search_channel_messages("quarterly review") - Find messages about quarterly reviews
        - search_channel_messages("john smith") - Find messages from/mentioning John Smith
    """
    logger.info(
        f"search_channel_messages called: query='{query}', team_id={team_id}, limit={limit}"
    )

    try:
        # Use Microsoft Graph search API to search across channel messages
        # If team_id is provided, we could filter results, but Graph search doesn't directly support this
        # So we'll search all and filter afterward if needed
        result = list(graph.search_query(query, ["chatMessage"], limit))

        # Filter for channel messages only (not chat messages) and optionally by team
        channel_messages = []
        for message in result:
            # Channel messages have teamId and channelId, chat messages have chatId
            if "teamId" in message and "channelId" in message:
                if team_id is None or message.get("teamId") == team_id:
                    try:
                        # Get team and channel context
                        t_id = message["teamId"]
                        c_id = message["channelId"]

                        team_info = graph.request("GET", f"/teams/{t_id}")
                        channel_info = graph.request(
                            "GET", f"/teams/{t_id}/channels/{c_id}"
                        )

                        if team_info and channel_info:
                            message["teamName"] = team_info.get("displayName", "")
                            message["channelName"] = channel_info.get("displayName", "")
                            message["webUrl"] = channel_info.get("webUrl", "")
                    except Exception as context_error:
                        logger.warning(
                            f"Could not get channel context: {str(context_error)}"
                        )

                    # Convert HTML to markdown if needed
                    if "body" in message and "content" in message["body"]:
                        if message["body"].get("contentType") == "html":
                            message["body"]["content"] = convert_to_markdown(
                                message["body"]["content"]
                            )
                            message["body"]["contentType"] = "text/markdown"

                    channel_messages.append(message)

        # Limit results
        result = channel_messages[:limit]

        logger.info(
            f"search_channel_messages successful: found {len(result)} channel messages matching '{query}'"
            + (f" in team {team_id}" if team_id else "")
        )
        return result
    except Exception as e:
        logger.error(
            f"search_channel_messages failed for query='{query}', team_id={team_id}: {str(e)}",
            exc_info=True,
        )
        raise
