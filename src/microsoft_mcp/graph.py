import httpx
import time
from typing import Any, Iterator, Optional
from .auth import AzureAuthentication

BASE_URL = "https://graph.microsoft.com/v1.0"
# 15 x 320 KiB = 4,915,200 bytes
UPLOAD_CHUNK_SIZE = 15 * 320 * 1024

_client = httpx.Client(timeout=30.0, follow_redirects=True)

# Global auth instance
_global_auth: Optional[AzureAuthentication] = None


def set_auth_instance(auth: AzureAuthentication) -> None:
    """Set the global authentication instance for the graph module"""
    global _global_auth
    _global_auth = auth


def get_auth_instance() -> AzureAuthentication:
    """Get the global authentication instance, creating one if needed"""
    global _global_auth
    if _global_auth is None:
        _global_auth = AzureAuthentication()
    return _global_auth


def request(
    method: str,
    path: str,
    params: dict[str, Any] | None = None,
    json: dict[str, Any] | None = None,
    data: bytes | None = None,
    max_retries: int = 3,
    auth: Optional[AzureAuthentication] = None,
) -> dict[str, Any] | None:
    auth_instance = auth or get_auth_instance()
    headers = {
        "Authorization": f"Bearer {auth_instance.get_token()}",
    }

    if method == "GET":
        if "$search" in (params or {}):
            headers["Prefer"] = 'outlook.body-content-type="text"'
        elif "body" in (params or {}).get("$select", ""):
            headers["Prefer"] = 'outlook.body-content-type="text"'
    else:
        headers["Content-Type"] = (
            "application/json" if json else "application/octet-stream"
        )

    if params and (
        "$search" in params
        or "contains(" in params.get("$filter", "")
        or "/any(" in params.get("$filter", "")
    ):
        headers["ConsistencyLevel"] = "eventual"
        params.setdefault("$count", "true")

    retry_count = 0
    while retry_count <= max_retries:
        try:
            response = _client.request(
                method=method,
                url=f"{BASE_URL}{path}",
                headers=headers,
                params=params,
                json=json,
                content=data,
            )

            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", "5"))
                if retry_count < max_retries:
                    time.sleep(min(retry_after, 60))
                    retry_count += 1
                    continue

            if response.status_code >= 500 and retry_count < max_retries:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue

            response.raise_for_status()

            if response.content:
                return response.json()
            return None

        except httpx.HTTPStatusError as e:
            if retry_count < max_retries and e.response.status_code >= 500:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue
            raise

    return None


def request_paginated(
    path: str,
    params: dict[str, Any] | None = None,
    limit: int | None = None,
    auth: Optional[AzureAuthentication] = None,
) -> Iterator[dict[str, Any]]:
    """Make paginated requests following @odata.nextLink"""
    items_returned = 0
    next_link = None

    while True:
        if next_link:
            result = request("GET", next_link.replace(BASE_URL, ""), auth=auth)
        else:
            result = request("GET", path, params=params, auth=auth)

        if not result:
            break

        if "value" in result:
            for item in result["value"]:
                if limit and items_returned >= limit:
                    return
                yield item
                items_returned += 1

        next_link = result.get("@odata.nextLink")
        if not next_link:
            break


def download_raw(
    path: str, max_retries: int = 3, auth: Optional[AzureAuthentication] = None
) -> bytes:
    auth_instance = auth or get_auth_instance()
    headers = {"Authorization": f"Bearer {auth_instance.get_token()}"}

    retry_count = 0
    while retry_count <= max_retries:
        try:
            response = _client.get(f"{BASE_URL}{path}", headers=headers)

            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", "5"))
                if retry_count < max_retries:
                    time.sleep(min(retry_after, 60))
                    retry_count += 1
                    continue

            if response.status_code >= 500 and retry_count < max_retries:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue

            response.raise_for_status()
            return response.content

        except httpx.HTTPStatusError as e:
            if retry_count < max_retries and e.response.status_code >= 500:
                wait_time = (2**retry_count) * 1
                time.sleep(wait_time)
                retry_count += 1
                continue
            raise

    raise ValueError("Failed to download file after all retries")


def _do_chunked_upload(
    upload_url: str,
    data: bytes,
    headers: dict[str, str],
) -> dict[str, Any]:
    """Internal helper for chunked uploads"""
    file_size = len(data)

    for i in range(0, file_size, UPLOAD_CHUNK_SIZE):
        chunk_start = i
        chunk_end = min(i + UPLOAD_CHUNK_SIZE, file_size)
        chunk = data[chunk_start:chunk_end]

        chunk_headers = headers.copy()
        chunk_headers["Content-Length"] = str(len(chunk))
        chunk_headers["Content-Range"] = (
            f"bytes {chunk_start}-{chunk_end - 1}/{file_size}"
        )

        retry_count = 0
        while retry_count <= 3:
            try:
                response = _client.put(upload_url, content=chunk, headers=chunk_headers)

                if response.status_code == 429:
                    retry_after = int(response.headers.get("Retry-After", "5"))
                    if retry_count < 3:
                        time.sleep(min(retry_after, 60))
                        retry_count += 1
                        continue

                response.raise_for_status()

                if response.status_code in (200, 201):
                    return response.json()
                break

            except httpx.HTTPStatusError as e:
                if retry_count < 3 and e.response.status_code >= 500:
                    time.sleep((2**retry_count) * 1)
                    retry_count += 1
                    continue
                raise

    raise ValueError("Upload completed but no final response received")


def create_upload_session(
    path: str,
    item_properties: dict[str, Any] | None = None,
    auth: Optional[AzureAuthentication] = None,
) -> dict[str, Any]:
    """Create an upload session for large files"""
    payload = {"item": item_properties or {}}
    result = request("POST", f"{path}/createUploadSession", json=payload, auth=auth)
    if not result:
        raise ValueError("Failed to create upload session")
    return result


def upload_large_file(
    path: str,
    data: bytes,
    item_properties: dict[str, Any] | None = None,
    auth: Optional[AzureAuthentication] = None,
) -> dict[str, Any]:
    """Upload a large file using upload sessions"""
    file_size = len(data)

    if file_size <= UPLOAD_CHUNK_SIZE:
        result = request("PUT", f"{path}/content", data=data, auth=auth)
        if not result:
            raise ValueError("Failed to upload file")
        return result

    session = create_upload_session(path, item_properties, auth=auth)
    upload_url = session["uploadUrl"]

    auth_instance = auth or get_auth_instance()
    headers = {"Authorization": f"Bearer {auth_instance.get_token()}"}
    return _do_chunked_upload(upload_url, data, headers)


def create_mail_upload_session(
    message_id: str,
    attachment_item: dict[str, Any],
    auth: Optional[AzureAuthentication] = None,
) -> dict[str, Any]:
    """Create an upload session for large mail attachments"""
    result = request(
        "POST",
        f"/me/messages/{message_id}/attachments/createUploadSession",
        json={"AttachmentItem": attachment_item},
        auth=auth,
    )
    if not result:
        raise ValueError("Failed to create mail attachment upload session")
    return result


def upload_large_mail_attachment(
    message_id: str,
    name: str,
    data: bytes,
    content_type: str = "application/octet-stream",
    auth: Optional[AzureAuthentication] = None,
) -> dict[str, Any]:
    """Upload a large mail attachment using upload sessions"""
    file_size = len(data)

    attachment_item = {
        "attachmentType": "file",
        "name": name,
        "size": file_size,
        "contentType": content_type,
    }

    session = create_mail_upload_session(message_id, attachment_item, auth=auth)
    upload_url = session["uploadUrl"]

    auth_instance = auth or get_auth_instance()
    headers = {"Authorization": f"Bearer {auth_instance.get_token()}"}
    return _do_chunked_upload(upload_url, data, headers)


def search_query(
    query: str,
    entity_types: list[str],
    limit: int = 50,
    fields: list[str] | None = None,
    auth: Optional[AzureAuthentication] = None,
) -> Iterator[dict[str, Any]]:
    """Use the modern /search/query API endpoint"""
    # Validate entity types - Microsoft Graph search has specific requirements
    valid_entity_types = {
        "message",
        "event",
        "driveItem",
        # "list",
        # "listItem",
        # "site",
        "drive",
        "chatMessage",
        "person",
        # "externalItem",
    }

    # Filter to only valid entity types
    filtered_entity_types = [et for et in entity_types if et in valid_entity_types]

    if not filtered_entity_types:
        # If no valid entity types, return empty iterator
        return iter([])

    payload = {
        "requests": [
            {
                "entityTypes": filtered_entity_types,
                "query": {"queryString": query},
                "size": min(limit, 25),
                "from": 0,
            }
        ]
    }

    # Add fields if specified
    if fields:
        payload["requests"][0]["fields"] = fields

    # Add stored fields for better results
    payload["requests"][0]["storedFields"] = [
        "id",
        "name",
        "subject",
        "body",
        "from",
        "to",
        "receivedDateTime",
        "lastModifiedDateTime",
        "size",
        "contentType",
    ]

    items_returned = 0

    while True:
        try:
            result = request("POST", "/search/query", json=payload, auth=auth)

            if not result or "value" not in result:
                break

            for response in result["value"]:
                if "hitsContainers" in response:
                    for container in response["hitsContainers"]:
                        if "hits" in container:
                            for hit in container["hits"]:
                                if limit and items_returned >= limit:
                                    return
                                yield hit["resource"]
                                items_returned += 1

            # Check for more results
            has_more = False
            for response in result.get("value", []):
                for container in response.get("hitsContainers", []):
                    if container.get("moreResultsAvailable"):
                        has_more = True
                        break

            if not has_more:
                break

            # Update from parameter for next batch
            payload["requests"][0]["from"] += payload["requests"][0]["size"]

        except Exception as e:
            # Log the error and break to avoid infinite loops
            print(f"Search query error: {e}")
            break
