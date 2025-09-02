import base64
import datetime as dt
import logging
import pathlib as pl
from typing import Any
from fastmcp import FastMCP
from . import graph, auth

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

mcp = FastMCP("microsoft-mcp")

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


@mcp.tool
def list_emails(
    folder: str = "inbox",
    limit: int = 10,
    include_body: bool = True,
) -> list[dict[str, Any]]:
    """List emails from specified folder"""
    logger.info(f"list_emails called: folder={folder}, limit={limit}, include_body={include_body}")
    
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

        emails = list(
            graph.request_paginated(
                f"/me/mailFolders/{folder_path}/messages",
                params=params,
                limit=limit,
            )
        )

        logger.info(f"list_emails successful: retrieved {len(emails)} emails from folder {folder}")
        return emails
    except Exception as e:
        logger.error(f"list_emails failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_email(
    email_id: str,
    include_body: bool = True,
    body_max_length: int = 50000,
    include_attachments: bool = True,
) -> dict[str, Any]:
    """Get email details with size limits

    Args:
        email_id: The email ID
        include_body: Whether to include the email body (default: True)
        body_max_length: Maximum characters for body content (default: 50000)
        include_attachments: Whether to include attachment metadata (default: True)
    """
    logger.info(f"get_email called: email_id={email_id}, include_body={include_body}, body_max_length={body_max_length}, include_attachments={include_attachments}")
    
    try:
        params = {}
        if include_attachments:
            params["$expand"] = "attachments($select=id,name,size,contentType)"

        result = graph.request("GET", f"/me/messages/{email_id}", params=params)
        if not result:
            logger.error(f"get_email failed: Email with ID {email_id} not found")
            raise ValueError(f"Email with ID {email_id} not found")

        # Truncate body if needed
        if include_body and "body" in result and "content" in result["body"]:
            content = result["body"]["content"]
            if len(content) > body_max_length:
                result["body"]["content"] = (
                    content[:body_max_length]
                    + f"\n\n[Content truncated - {len(content)} total characters]"
                )
                result["body"]["truncated"] = True
                result["body"]["total_length"] = len(content)
                logger.info(f"get_email: body truncated from {len(content)} to {body_max_length} characters")
        elif not include_body and "body" in result:
            del result["body"]

        # Remove attachment content bytes to reduce size
        if "attachments" in result and result["attachments"]:
            for attachment in result["attachments"]:
                if "contentBytes" in attachment:
                    del attachment["contentBytes"]

        logger.info(f"get_email successful: retrieved email {email_id}")
        return result
    except Exception as e:
        logger.error(f"get_email failed for email_id={email_id}: {str(e)}", exc_info=True)
        raise


# @mcp.tool
# def create_email_draft(
#     to: str | list[str],
#     subject: str,
#     body: str,
#     cc: str | list[str] | None = None,
#     attachments: str | list[str] | None = None,
# ) -> dict[str, Any]:
#     """Create an email draft with file path(s) as attachments"""
#     logger.info(f"create_email_draft called: to={to}, subject='{subject}', cc={cc}, attachments={attachments}")
    
#     try:
#         to_list = [to] if isinstance(to, str) else to

#         message = {
#             "subject": subject,
#             "body": {"contentType": "Text", "content": body},
#             "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
#         }

#         if cc:
#             cc_list = [cc] if isinstance(cc, str) else cc
#             message["ccRecipients"] = [
#                 {"emailAddress": {"address": addr}} for addr in cc_list
#             ]

#         small_attachments = []
#         large_attachments = []

#         if attachments:
#             # Convert single path to list
#             attachment_paths = (
#                 [attachments] if isinstance(attachments, str) else attachments
#             )
#             logger.info(f"create_email_draft: processing {len(attachment_paths)} attachments")
            
#             for file_path in attachment_paths:
#                 try:
#                     path = pl.Path(file_path).expanduser().resolve()
#                     content_bytes = path.read_bytes()
#                     att_size = len(content_bytes)
#                     att_name = path.name

#                     if att_size < 3 * 1024 * 1024:
#                         small_attachments.append(
#                             {
#                                 "@odata.type": "#microsoft.graph.fileAttachment",
#                                 "name": att_name,
#                                 "contentBytes": base64.b64encode(content_bytes).decode("utf-8"),
#                             }
#                         )
#                         logger.info(f"create_email_draft: added small attachment {att_name} ({att_size} bytes)")
#                     else:
#                         large_attachments.append(
#                             {
#                                 "name": att_name,
#                                 "content_bytes": content_bytes,
#                                 "content_type": "application/octet-stream",
#                             }
#                         )
#                         logger.info(f"create_email_draft: queued large attachment {att_name} ({att_size} bytes)")
#                 except Exception as e:
#                     logger.error(f"create_email_draft: failed to process attachment {file_path}: {str(e)}")
#                     raise

#         if small_attachments:
#             message["attachments"] = small_attachments

#         result = graph.request("POST", "/me/messages", json=message)
#         if not result:
#             logger.error("create_email_draft failed: no response from server")
#             raise ValueError("Failed to create email draft")

#         message_id = result["id"]
#         logger.info(f"create_email_draft: created draft with ID {message_id}")

#         for att in large_attachments:
#             try:
#                 graph.upload_large_mail_attachment(
#                     message_id,
#                     att["name"],
#                     att["content_bytes"],
#                     att.get("content_type", "application/octet-stream"),
#                 )
#                 logger.info(f"create_email_draft: uploaded large attachment {att['name']}")
#             except Exception as e:
#                 logger.error(f"create_email_draft: failed to upload large attachment {att['name']}: {str(e)}")
#                 raise

#         logger.info(f"create_email_draft successful: created draft with {len(small_attachments)} small and {len(large_attachments)} large attachments")
#         return result
#     except Exception as e:
#         logger.error(f"create_email_draft failed: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def send_email(
#     to: str | list[str],
#     subject: str,
#     body: str,
#     cc: str | list[str] | None = None,
#     attachments: str | list[str] | None = None,
# ) -> dict[str, str]:
#     """Send an email immediately with file path(s) as attachments"""
#     logger.info(f"send_email called: to={to}, subject='{subject}', cc={cc}, attachments={attachments}")
    
#     try:
#         to_list = [to] if isinstance(to, str) else to

#         message = {
#             "subject": subject,
#             "body": {"contentType": "Text", "content": body},
#             "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
#         }

#         if cc:
#             cc_list = [cc] if isinstance(cc, str) else cc
#             message["ccRecipients"] = [
#                 {"emailAddress": {"address": addr}} for addr in cc_list
#             ]

#         # Check if we have large attachments
#         has_large_attachments = False
#         processed_attachments = []

#         if attachments:
#             # Convert single path to list
#             attachment_paths = (
#                 [attachments] if isinstance(attachments, str) else attachments
#             )
#             logger.info(f"send_email: processing {len(attachment_paths)} attachments")
            
#             for file_path in attachment_paths:
#                 try:
#                     path = pl.Path(file_path).expanduser().resolve()
#                     content_bytes = path.read_bytes()
#                     att_size = len(content_bytes)
#                     att_name = path.name

#                     processed_attachments.append(
#                         {
#                             "name": att_name,
#                             "content_bytes": content_bytes,
#                             "content_type": "application/octet-stream",
#                             "size": att_size,
#                         }
#                     )

#                     if att_size >= 3 * 1024 * 1024:
#                         has_large_attachments = True
#                         logger.info(f"send_email: detected large attachment {att_name} ({att_size} bytes)")
#                     else:
#                         logger.info(f"send_email: processed small attachment {att_name} ({att_size} bytes)")
#                 except Exception as e:
#                     logger.error(f"send_email: failed to process attachment {file_path}: {str(e)}")
#                     raise

#         if not has_large_attachments and processed_attachments:
#             message["attachments"] = [
#                 {
#                     "@odata.type": "#microsoft.graph.fileAttachment",
#                     "name": att["name"],
#                     "contentBytes": base64.b64encode(att["content_bytes"]).decode("utf-8"),
#                 }
#                 for att in processed_attachments
#             ]
#             graph.request("POST", "/me/sendMail", json={"message": message})
#             logger.info(f"send_email successful: sent email with {len(processed_attachments)} small attachments")
#             return {"status": "sent"}
#         elif has_large_attachments:
#             # Create draft first, then add large attachments, then send
#             logger.info("send_email: handling large attachments via draft method")
#             to_list = [to] if isinstance(to, str) else to
#             message = {
#                 "subject": subject,
#                 "body": {"contentType": "Text", "content": body},
#                 "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
#             }
#             if cc:
#                 cc_list = [cc] if isinstance(cc, str) else cc
#                 message["ccRecipients"] = [
#                     {"emailAddress": {"address": addr}} for addr in cc_list
#                 ]

#             result = graph.request("POST", "/me/messages", json=message)
#             if not result:
#                 logger.error("send_email failed: could not create draft for large attachments")
#                 raise ValueError("Failed to create email draft")

#             message_id = result["id"]
#             logger.info(f"send_email: created draft {message_id} for large attachments")

#             for att in processed_attachments:
#                 try:
#                     if att["size"] >= 3 * 1024 * 1024:
#                         graph.upload_large_mail_attachment(
#                             message_id,
#                             att["name"],
#                             att["content_bytes"],
#                             att.get("content_type", "application/octet-stream"),
#                         )
#                         logger.info(f"send_email: uploaded large attachment {att['name']}")
#                     else:
#                         small_att = {
#                             "@odata.type": "#microsoft.graph.fileAttachment",
#                             "name": att["name"],
#                             "contentBytes": base64.b64encode(att["content_bytes"]).decode(
#                                 "utf-8"
#                             ),
#                         }
#                         graph.request(
#                             "POST",
#                             f"/me/messages/{message_id}/attachments",
#                             json=small_att,
#                         )
#                         logger.info(f"send_email: uploaded small attachment {att['name']}")
#                 except Exception as e:
#                     logger.error(f"send_email: failed to upload attachment {att['name']}: {str(e)}")
#                     raise

#             graph.request("POST", f"/me/messages/{message_id}/send")
#             logger.info(f"send_email successful: sent email with {len(processed_attachments)} attachments via draft method")
#             return {"status": "sent"}
#         else:
#             graph.request("POST", "/me/sendMail", json={"message": message})
#             logger.info("send_email successful: sent email without attachments")
#             return {"status": "sent"}
#     except Exception as e:
#         logger.error(f"send_email failed: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def update_email(email_id: str, updates: dict[str, Any]) -> dict[str, Any]:
#     """Update email properties (isRead, categories, flag, etc.)"""
#     logger.info(f"update_email called: email_id={email_id}, updates={updates}")
    
#     try:
#         result = graph.request("PATCH", f"/me/messages/{email_id}", json=updates)
#         if not result:
#             logger.error(f"update_email failed: no response for email {email_id}")
#             raise ValueError(f"Failed to update email {email_id} - no response")
        
#         logger.info(f"update_email successful: updated email {email_id}")
#         return result
#     except Exception as e:
#         logger.error(f"update_email failed for email_id={email_id}: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def delete_email(email_id: str) -> dict[str, str]:
#     """Delete an email"""
#     logger.info(f"delete_email called: email_id={email_id}")
    
#     try:
#         graph.request("DELETE", f"/me/messages/{email_id}")
#         logger.info(f"delete_email successful: deleted email {email_id}")
#         return {"status": "deleted"}
#     except Exception as e:
#         logger.error(f"delete_email failed for email_id={email_id}: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def move_email(email_id: str, destination_folder: str) -> dict[str, Any]:
#     """Move email to another folder"""
#     logger.info(f"move_email called: email_id={email_id}, destination_folder={destination_folder}")
    
#     try:
#         folder_path = FOLDERS.get(destination_folder.casefold(), destination_folder)

#         folders = graph.request("GET", "/me/mailFolders")
#         folder_id = None

#         if not folders:
#             logger.error("move_email failed: could not retrieve mail folders")
#             raise ValueError("Failed to retrieve mail folders")
#         if "value" not in folders:
#             logger.error(f"move_email failed: unexpected folder response structure: {folders}")
#             raise ValueError(f"Unexpected folder response structure: {folders}")

#         for folder in folders["value"]:
#             if folder["displayName"].lower() == folder_path.lower():
#                 folder_id = folder["id"]
#                 break

#         if not folder_id:
#             logger.error(f"move_email failed: folder '{destination_folder}' not found")
#             raise ValueError(f"Folder '{destination_folder}' not found")

#         payload = {"destinationId": folder_id}
#         result = graph.request("POST", f"/me/messages/{email_id}/move", json=payload)
#         if not result:
#             logger.error("move_email failed: no response from server")
#             raise ValueError("Failed to move email - no response from server")
#         if "id" not in result:
#             logger.error(f"move_email failed: unexpected response: {result}")
#             raise ValueError(f"Failed to move email - unexpected response: {result}")
        
#         logger.info(f"move_email successful: moved email {email_id} to folder {destination_folder}, new_id={result['id']}")
#         return {"status": "moved", "new_id": result["id"]}
#     except Exception as e:
#         logger.error(f"move_email failed for email_id={email_id}, destination_folder={destination_folder}: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def reply_to_email(email_id: str, body: str) -> dict[str, str]:
#     """Reply to an email (sender only)"""
#     logger.info(f"reply_to_email called: email_id={email_id}")
    
#     try:
#         endpoint = f"/me/messages/{email_id}/reply"
#         payload = {"message": {"body": {"contentType": "Text", "content": body}}}
#         graph.request("POST", endpoint, json=payload)
#         logger.info(f"reply_to_email successful: replied to email {email_id}")
#         return {"status": "sent"}
#     except Exception as e:
#         logger.error(f"reply_to_email failed for email_id={email_id}: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def reply_all_email(email_id: str, body: str) -> dict[str, str]:
#     """Reply to all recipients of an email"""
#     logger.info(f"reply_all_email called: email_id={email_id}")
    
#     try:
#         endpoint = f"/me/messages/{email_id}/replyAll"
#         payload = {"message": {"body": {"contentType": "Text", "content": body}}}
#         graph.request("POST", endpoint, json=payload)
#         logger.info(f"reply_all_email successful: replied to all for email {email_id}")
#         return {"status": "sent"}
#     except Exception as e:
#         logger.error(f"reply_all_email failed for email_id={email_id}: {str(e)}", exc_info=True)
#         raise


@mcp.tool
def list_events(
    days_ahead: int = 7,
    days_back: int = 0,
    include_details: bool = True,
) -> list[dict[str, Any]]:
    """List calendar events within specified date range, including recurring event instances"""
    logger.info(f"list_events called: days_ahead={days_ahead}, days_back={days_back}, include_details={include_details}")
    
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

        logger.info(f"list_events successful: retrieved {len(events)} events from {start} to {end}")
        return events
    except Exception as e:
        logger.error(f"list_events failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_event(event_id: str) -> dict[str, Any]:
    """Get full event details"""
    logger.info(f"get_event called: event_id={event_id}")
    
    try:
        result = graph.request("GET", f"/me/events/{event_id}")
        if not result:
            logger.error(f"get_event failed: Event with ID {event_id} not found")
            raise ValueError(f"Event with ID {event_id} not found")
        
        logger.info(f"get_event successful: retrieved event {event_id}")
        return result
    except Exception as e:
        logger.error(f"get_event failed for event_id={event_id}: {str(e)}", exc_info=True)
        raise


# @mcp.tool
# def create_event(
#     subject: str,
#     start: str,
#     end: str,
#     location: str | None = None,
#     body: str | None = None,
#     attendees: str | list[str] | None = None,
#     timezone: str = "UTC",
# ) -> dict[str, Any]:
#     """Create a calendar event"""
#     logger.info(f"create_event called: subject='{subject}', start={start}, end={end}, location={location}, attendees={attendees}, timezone={timezone}")
    
#     try:
#         event = {
#             "subject": subject,
#             "start": {"dateTime": start, "timeZone": timezone},
#             "end": {"dateTime": end, "timeZone": timezone},
#         }

#         if location:
#             event["location"] = {"displayName": location}

#         if body:
#             event["body"] = {"contentType": "Text", "content": body}

#         if attendees:
#             attendees_list = [attendees] if isinstance(attendees, str) else attendees
#             event["attendees"] = [
#                 {"emailAddress": {"address": a}, "type": "required"} for a in attendees_list
#             ]
#             logger.info(f"create_event: added {len(attendees_list)} attendees")

#         result = graph.request("POST", "/me/events", json=event)
#         if not result:
#             logger.error("create_event failed: no response from server")
#             raise ValueError("Failed to create event")
        
#         logger.info(f"create_event successful: created event with ID {result.get('id')}")
#         return result
#     except Exception as e:
#         logger.error(f"create_event failed: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def update_event(event_id: str, updates: dict[str, Any]) -> dict[str, Any]:
#     """Update event properties"""
#     logger.info(f"update_event called: event_id={event_id}, updates={updates}")
    
#     try:
#         formatted_updates = {}

#         if "subject" in updates:
#             formatted_updates["subject"] = updates["subject"]
#         if "start" in updates:
#             formatted_updates["start"] = {
#                 "dateTime": updates["start"],
#                 "timeZone": updates.get("timezone", "UTC"),
#             }
#         if "end" in updates:
#             formatted_updates["end"] = {
#                 "dateTime": updates["end"],
#                 "timeZone": updates.get("timezone", "UTC"),
#             }
#         if "location" in updates:
#             formatted_updates["location"] = {"displayName": updates["location"]}
#         if "body" in updates:
#             formatted_updates["body"] = {"contentType": "Text", "content": updates["body"]}

#         result = graph.request("PATCH", f"/me/events/{event_id}", json=formatted_updates)
#         logger.info(f"update_event successful: updated event {event_id}")
#         return result or {"status": "updated"}
#     except Exception as e:
#         logger.error(f"update_event failed for event_id={event_id}: {str(e)}", exc_info=True)
#         raise


# @mcp.tool
# def delete_event(event_id: str, send_cancellation: bool = True) -> dict[str, str]:
#     """Delete or cancel a calendar event"""
#     if send_cancellation:
#         graph.request("POST", f"/me/events/{event_id}/cancel", json={})
#     else:
#         graph.request("DELETE", f"/me/events/{event_id}")
#     return {"status": "deleted"}


# @mcp.tool
# def respond_event(
#     event_id: str,
#     response: str = "accept",
#     message: str | None = None,
# ) -> dict[str, str]:
#     """Respond to event invitation (accept, decline, tentativelyAccept)"""
#     payload: dict[str, Any] = {"sendResponse": True}
#     if message:
#         payload["comment"] = message

#     graph.request("POST", f"/me/events/{event_id}/{response}", json=payload)
#     return {"status": response}


@mcp.tool
def check_availability(
    start: str,
    end: str,
    attendees: str | list[str] | None = None,
) -> dict[str, Any]:
    """Check calendar availability for scheduling"""
    me_info = graph.request("GET", "/me")
    if not me_info or "mail" not in me_info:
        raise ValueError("Failed to get user email address")
    schedules = [me_info["mail"]]
    if attendees:
        attendees_list = [attendees] if isinstance(attendees, str) else attendees
        schedules.extend(attendees_list)

    payload = {
        "schedules": schedules,
        "startTime": {"dateTime": start, "timeZone": "UTC"},
        "endTime": {"dateTime": end, "timeZone": "UTC"},
        "availabilityViewInterval": 30,
    }

    result = graph.request("POST", "/me/calendar/getSchedule", json=payload)
    if not result:
        raise ValueError("Failed to check availability")
    return result


@mcp.tool
def list_contacts(limit: int = 50) -> list[dict[str, Any]]:
    """List contacts"""
    params = {"$top": min(limit, 100)}

    contacts = list(graph.request_paginated("/me/contacts", params=params, limit=limit))

    return contacts


@mcp.tool
def get_contact(contact_id: str) -> dict[str, Any]:
    """Get contact details"""
    result = graph.request("GET", f"/me/contacts/{contact_id}")
    if not result:
        raise ValueError(f"Contact with ID {contact_id} not found")
    return result


# @mcp.tool
# def create_contact(
#     given_name: str,
#     surname: str | None = None,
#     email_addresses: str | list[str] | None = None,
#     phone_numbers: dict[str, str] | None = None,
# ) -> dict[str, Any]:
#     """Create a new contact"""
#     contact: dict[str, Any] = {"givenName": given_name}

#     if surname:
#         contact["surname"] = surname

#     if email_addresses:
#         email_list = (
#             [email_addresses] if isinstance(email_addresses, str) else email_addresses
#         )
#         contact["emailAddresses"] = [
#             {"address": email, "name": f"{given_name} {surname or ''}".strip()}
#             for email in email_list
#         ]

#     if phone_numbers:
#         if "business" in phone_numbers:
#             contact["businessPhones"] = [phone_numbers["business"]]
#         if "home" in phone_numbers:
#             contact["homePhones"] = [phone_numbers["home"]]
#         if "mobile" in phone_numbers:
#             contact["mobilePhone"] = phone_numbers["mobile"]

#     result = graph.request("POST", "/me/contacts", json=contact)
#     if not result:
#         raise ValueError("Failed to create contact")
#     return result


# @mcp.tool
# def update_contact(contact_id: str, updates: dict[str, Any]) -> dict[str, Any]:
#     """Update contact information"""
#     result = graph.request("PATCH", f"/me/contacts/{contact_id}", json=updates)
#     return result or {"status": "updated"}


# @mcp.tool
# def delete_contact(contact_id: str) -> dict[str, str]:
#     """Delete a contact"""
#     graph.request("DELETE", f"/me/contacts/{contact_id}")
#     return {"status": "deleted"}


@mcp.tool
def list_files(path: str = "/", limit: int = 50) -> list[dict[str, Any]]:
    """List files and folders in OneDrive"""
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

    return [
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


@mcp.tool
def get_file(file_id: str, download_path: str) -> dict[str, Any]:
    """Download a file from OneDrive to local path"""
    import subprocess

    metadata = graph.request("GET", f"/me/drive/items/{file_id}")
    if not metadata:
        raise ValueError(f"File with ID {file_id} not found")

    download_url = metadata.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise ValueError("No download URL available for this file")

    try:
        subprocess.run(
            ["curl", "-L", "-o", download_path, download_url],
            check=True,
            capture_output=True,
        )

        return {
            "path": download_path,
            "name": metadata.get("name", "unknown"),
            "size_mb": round(metadata.get("size", 0) / (1024 * 1024), 2),
            "mime_type": metadata.get("file", {}).get("mimeType") if metadata else None,
        }
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Failed to download file: {e.stderr.decode()}")


# @mcp.tool
# def create_file(onedrive_path: str, local_file_path: str) -> dict[str, Any]:
#     """Upload a local file to OneDrive"""
#     path = pl.Path(local_file_path).expanduser().resolve()
#     data = path.read_bytes()
#     result = graph.upload_large_file(f"/me/drive/root:/{onedrive_path}:", data)
#     if not result:
#         raise ValueError(f"Failed to create file at path: {onedrive_path}")
#     return result


# @mcp.tool
# def update_file(file_id: str, local_file_path: str) -> dict[str, Any]:
#     """Update OneDrive file content from a local file"""
#     path = pl.Path(local_file_path).expanduser().resolve()
#     data = path.read_bytes()
#     result = graph.upload_large_file(f"/me/drive/items/{file_id}", data)
#     if not result:
#         raise ValueError(f"Failed to update file with ID: {file_id}")
#     return result


# @mcp.tool
# def delete_file(file_id: str) -> dict[str, str]:
#     """Delete a file or folder"""
#     graph.request("DELETE", f"/me/drive/items/{file_id}")
#     return {"status": "deleted"}


@mcp.tool
def get_attachment(email_id: str, attachment_id: str, save_path: str) -> dict[str, Any]:
    """Download email attachment to a specified file path"""
    result = graph.request(
        "GET", f"/me/messages/{email_id}/attachments/{attachment_id}"
    )

    if not result:
        raise ValueError("Attachment not found")

    if "contentBytes" not in result:
        raise ValueError("Attachment content not available")

    # Save attachment to file
    path = pl.Path(save_path).expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    content_bytes = base64.b64decode(result["contentBytes"])
    path.write_bytes(content_bytes)

    return {
        "name": result.get("name", "unknown"),
        "content_type": result.get("contentType", "application/octet-stream"),
        "size": result.get("size", 0),
        "saved_to": str(path),
    }


@mcp.tool
def search_files(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for files in OneDrive using the modern search API."""
    items = list(graph.search_query(query, ["driveItem"], limit))

    return [
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


@mcp.tool
def search_emails(
    query: str,
    limit: int = 50,
    folder: str | None = None,
) -> list[dict[str, Any]]:
    """Search emails using the modern search API."""
    if folder:
        # For folder-specific search, use the traditional endpoint
        folder_path = FOLDERS.get(folder.casefold(), folder)
        endpoint = f"/me/mailFolders/{folder_path}/messages"

        params = {
            "$search": f'"{query}"',
            "$top": min(limit, 100),
            "$select": "id,subject,from,toRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead",
        }

        return list(graph.request_paginated(endpoint, params=params, limit=limit))

    return list(graph.search_query(query, ["message"], limit))


@mcp.tool
def search_events(
    query: str,
    #days_ahead: int = 365,
    #days_back: int = 365,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search calendar events using the modern search API."""
    events = list(graph.search_query(query, ["event"], limit))

    # # Filter by date range if needed
    # if days_ahead != 365 or days_back != 365:
    #     now = dt.datetime.now(dt.timezone.utc)
    #     start = now - dt.timedelta(days=days_back)
    #     end = now + dt.timedelta(days=days_ahead)

    #     filtered_events = []
    #     for event in events:
    #         event_start = dt.datetime.fromisoformat(
    #             event.get("start", {}).get("dateTime", "").replace("Z", "+00:00")
    #         )
    #         event_end = dt.datetime.fromisoformat(
    #             event.get("end", {}).get("dateTime", "").replace("Z", "+00:00")
    #         )

    #         if event_start <= end and event_end >= start:
    #             filtered_events.append(event)

    #     return filtered_events

    return events


@mcp.tool
def search_contacts(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search contacts. Uses traditional search since unified_search doesn't support contacts."""
    params = {
        "$search": f'"{query}"',
        "$top": min(limit, 100),
    }

    contacts = list(graph.request_paginated("/me/contacts", params=params, limit=limit))

    return contacts


# @mcp.tool
# def unified_search(
#     query: str,
#     entity_types: list[str] | None = None,
#     limit: int = 50,
# ) -> dict[str, list[dict[str, Any]]]:
#     """Search across multiple Microsoft 365 resources using the modern search API

#     entity_types can include: 'message', 'event', 'drive', 'driveItem', 'chatMessage', 'person'
#     If not specified, searches across all available types.
#     """
#     if not entity_types:
#         entity_types = ["message", "chatMessage", "event"]

#     results = {entity_type: [] for entity_type in entity_types}

#     try:
#         items = list(graph.search_query(query, entity_types, limit))

#         for item in items:
#             resource_type = item.get("@odata.type", "").split(".")[-1]

#             if resource_type == "message":
#                 results.setdefault("message", []).append(item)
#             elif resource_type == "event":
#                 results.setdefault("event", []).append(item)
#             elif resource_type in ["driveItem", "file", "folder"]:
#                 results.setdefault("driveItem", []).append(item)
#             else:
#                 results.setdefault("other", []).append(item)

#         return {k: v for k, v in results.items() if v}

#     except Exception as e:
#         # Return error information in a structured format
#         return {
#             "error": [
#                 {
#                     "message": f"Search failed: {str(e)}",
#                     "entity_types": entity_types,
#                     "query": query,
#                 }
#             ]
#         }
