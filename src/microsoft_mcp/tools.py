import base64
import datetime as dt
import logging
import pathlib as pl
import subprocess
from typing import Any
from unittest import result
from urllib.parse import quote
from fastmcp import FastMCP
from . import graph
from .auth import AzureAuthentication
from markitdown import MarkItDown, StreamInfo
from io import BytesIO

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

mcp = FastMCP("microsoft-graph-mcp")

# Create a global authentication instance
auth = AzureAuthentication()

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


@mcp.tool
def is_logged_in() -> bool:
    return auth.exists_valid_token()


@mcp.tool
def login() -> str:
    """Ensure the user is authenticated and return user info.
    Raises an error if authentication fails.

    `login` is required before any other tools can succeed.
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
                email["conversation_url"] = f"https://outlook.office.com/mail/deeplink/readconv/{quote(email['conversationId'])}"


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
            result["conversation_url"] = f"https://outlook.office.com/mail/deeplink/readconv/{quote(result['conversationId'])}"

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
    include_details: bool = True,
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
        - list_events(include_details=False) - Get basic event info only for faster response
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
        - availabilityView: Numeric representation of availability (0=free, 1=tentative, 2=busy)

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


@mcp.tool
def search_files(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for files and folders in OneDrive using text queries.

    Find files by name, content, or metadata across your entire OneDrive. More powerful than
    browsing folders - searches filenames, document content, and file properties.

    Args:
        query: Search terms to find files (e.g., "budget report", "vacation photos", "presentation")
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching file/folder objects containing:
        - Basic info: id, name, type (file/folder), size (bytes), modified (timestamp)
        - Download info: download_url for direct file access (for files only)
        - Results ranked by relevance to search query

    Examples:
        - search_files("presentation") - Find files with "presentation" in name or content
        - search_files("budget 2024") - Find budget-related files from 2024
        - search_files("photos vacation") - Find vacation photos
        - search_files(".pdf report") - Find PDF files containing "report"
    """
    logger.info(f"search_files called: query='{query}', limit={limit}")

    try:
        items = list(graph.search_query(query, ["driveItem"], limit))

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
            f"search_files successful: found {len(result)} files matching '{query}'"
        )
        return result
    except Exception as e:
        logger.error(
            f"search_files failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


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
                    email["conversation_url"] = f"https://outlook.office.com/mail/deeplink/readconv/{quote(email['conversationId'])}"
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
                email["conversation_url"] = f"https://outlook.office.com/mail/deeplink/readconv/{quote(email['conversationId'])}"
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
