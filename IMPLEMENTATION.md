# Implementation Overview

This document describes the key implementation concepts and architecture of the Microsoft MCP (Model Context Protocol) server for Microsoft Graph API integration.

## Project Overview

Microsoft MCP is a comprehensive MCP server that provides AI assistants with seamless access to Microsoft 365 services including Outlook (Email), Calendar, OneDrive (Files), and Contacts through the Microsoft Graph API. The implementation focuses on delegated access authentication, allowing the application to act on behalf of signed-in users.

## Architecture

### Core Components

#### 1. Authentication System (`auth.py`)
- **Delegated Access**: Uses Azure Identity's `InteractiveBrowserCredential` for user authentication
- **Modern Authentication Flow**: Implements authorization code flow with PKCE (Proof Key for Code Exchange)
- **Token Caching**: Two-level caching system:
  - Azure Identity automatic caching
  - Application-level token cache (`~/.microsoft_mcp_delegated_token_cache.json`)
- **Scope Management**: Requests specific delegated permissions rather than broad access
- **Browser-based Auth**: Opens browser for user sign-in, no device codes required

**Key Features:**
- Automatic token refresh
- 5-minute expiration buffer for token validity
- Cache validation and cleanup
- Support for multiple tenants (common, consumers, organization-specific)

#### 2. Graph API Client (`graph.py`)
- **HTTP Client**: Uses `httpx` for robust HTTP communication
- **Retry Logic**: Implements exponential backoff for rate limiting (429) and server errors (5xx)
- **Pagination Support**: Handles Microsoft Graph `@odata.nextLink` pagination automatically
- **Large File Uploads**: Chunked upload sessions for files >3MB (emails) or custom chunk sizes (OneDrive)
- **Search Integration**: Modern `/search/query` API endpoint support

**Key Features:**
- Request/response logging
- Automatic header management (Authorization, Content-Type, ConsistencyLevel)
- Upload session management for large attachments
- Download capabilities with streaming support

#### 3. MCP Tools (`tools.py`)
- **FastMCP Framework**: Uses FastMCP for tool registration and management
- **Comprehensive Coverage**: 20+ tools covering email, calendar, contacts, and files
- **Error Handling**: Consistent error logging and exception propagation
- **Response Optimization**: Configurable body truncation, attachment handling

#### 4. Server Implementation (`server.py`)
- **Environment Validation**: Checks for required `MICROSOFT_MCP_CLIENT_ID`
- **Startup Authentication**: Validates authentication before starting MCP server
- **Error Recovery**: Graceful failure with helpful error messages

## Implementation Patterns

### 1. Delegated Access Model
The system implements delegated access where the application acts on behalf of the authenticated user rather than with its own identity. This provides:
- User-scoped data access
- Respect for user permissions
- No need for administrative consent in most cases
- Secure token management

### 2. Error Handling Strategy
```python
try:
    # Operation
    result = graph.request(...)
    logger.info(f"Operation successful: {details}")
    return result
except Exception as e:
    logger.error(f"Operation failed: {str(e)}", exc_info=True)
    raise
```

### 3. Pagination Pattern
```python
def request_paginated(path, params=None, limit=None):
    items_returned = 0
    next_link = None
    
    while True:
        result = request("GET", next_link or path, params=params)
        for item in result.get("value", []):
            if limit and items_returned >= limit:
                return
            yield item
            items_returned += 1
        
        next_link = result.get("@odata.nextLink")
        if not next_link:
            break
```

### 4. Large File Handling
- **Email Attachments**: 3MB threshold for chunked uploads
- **OneDrive Files**: Configurable chunk size (15 x 320KB = ~5MB chunks)
- **Upload Sessions**: Create session → Upload chunks → Finalize

## Tool Categories

### Email Tools (9 tools)
- **Core Operations**: list, get, create_draft, send, reply, reply_all
- **Management**: update, move, delete
- **Search**: search_emails, get_attachment
- **Features**: Attachment support, folder management, thread handling

### Calendar Tools (7 tools)
- **Core Operations**: list_events, get_event, create_event, update_event, delete_event
- **Interaction**: respond_event, check_availability
- **Features**: Recurring events, attendee management, availability checking

### Contact Tools (6 tools)
- **Core Operations**: list_contacts, get_contact, create_contact, update_contact, delete_contact
- **Search**: search_contacts
- **Features**: Multiple email addresses, phone numbers, addresses

### File Tools (6 tools)
- **Core Operations**: list_files, get_file, create_file, update_file, delete_file
- **Search**: search_files
- **Features**: Path-based navigation, download/upload, metadata management

### Utility Tools (1 tool)
- **unified_search**: Cross-service search across emails, events, files

## Configuration

### Environment Variables
- `MICROSOFT_MCP_CLIENT_ID`: Azure AD application ID (required)
- `MICROSOFT_MCP_TENANT_ID`: Tenant ID (optional, defaults to "common")
- `MICROSOFT_MCP_REDIRECT_URI`: Custom redirect URI (optional, for non-localhost deployments)

### Required Azure Permissions
```python
SCOPES = [
    "User.Read",                    # Read user profile
    "User.ReadBasic.All",          # Read basic user info
    "Chat.Read",                   # Read chat messages
    "Mail.Read",                   # Read emails
    "Team.ReadBasic.All",          # Read team info
    "TeamMember.ReadWrite.All",    # Manage team membership
    "Calendars.Read",              # Read calendars
    "Files.Read",                  # Read OneDrive files
]
```

## Key Design Decisions

### 1. Delegated vs Application Access
- **Chosen**: Delegated access
- **Rationale**: User-scoped permissions, better security model, no admin consent required
- **Trade-off**: Requires user authentication vs automatic background access

### 2. Token Caching Strategy
- **Two-level caching**: Azure Identity + application cache
- **Benefits**: Reduced authentication prompts, improved performance
- **Implementation**: File-based cache with expiration validation

### 3. Error Handling Philosophy
- **Fail Fast**: Validate inputs early, provide clear error messages
- **Logging**: Comprehensive logging for debugging
- **User Experience**: Helpful error messages, recovery suggestions

### 4. Response Size Management
- **Body Truncation**: Configurable limits for email body content
- **Attachment Handling**: Metadata only unless explicitly requested
- **Pagination**: Limit-based result sets to manage response sizes

## Security Considerations

### 1. Token Security
- Local file-based cache (`~/.microsoft_mcp_delegated_token_cache.json`)
- Tokens have expiration times
- Cache can be cleared manually
- No tokens in environment variables or code

### 2. Permission Model
- Principle of least privilege
- Specific scopes requested
- User consent required
- Delegated (not application) permissions

### 3. Data Handling
- No persistent data storage
- Temporary files for downloads/uploads
- Memory-efficient streaming for large files

## Development and Testing

### Project Structure
```
src/microsoft_mcp/
├── __init__.py          # Package initialization
├── server.py            # MCP server entry point
├── auth.py              # Authentication and token management
├── graph.py             # Microsoft Graph API client
├── tools.py             # MCP tool implementations
└── tools_old.py         # Legacy tools (reference)

tests/
└── test_integration.py  # Comprehensive integration tests

authenticate.py          # Standalone authentication script
```

### Testing Strategy
- **Integration Tests**: Full end-to-end testing with real Microsoft Graph API
- **Authentication Testing**: Multi-account scenarios
- **Tool Coverage**: All tools tested with real data
- **Error Scenarios**: Network failures, invalid inputs, permission issues

### Development Workflow
1. Environment setup with Azure app registration
2. Authentication using `authenticate.py`
3. Development with uv/Python tooling
4. Testing with pytest
5. Code formatting with black/ruff

## Performance Characteristics

### Typical Response Times
- **Token acquisition**: 100-500ms (cached) / 2-5s (fresh auth)
- **Simple API calls**: 200-800ms
- **Paginated requests**: 500ms-2s per page
- **File uploads**: Depends on size, ~1MB/s
- **Search operations**: 800ms-2s

### Rate Limiting
- Microsoft Graph: ~1000 requests/minute/tenant
- Automatic retry with exponential backoff
- 429 status code handling with Retry-After headers

### Memory Usage
- Minimal memory footprint
- Streaming for large files
- No persistent caching in memory
- HTTP connection pooling via httpx

## Future Considerations

### Potential Enhancements
1. **Multi-account support**: Manage multiple Microsoft accounts simultaneously
2. **Webhook subscriptions**: Real-time notifications for changes
3. **Batch operations**: Multiple API calls in single request
4. **Advanced search**: More sophisticated query capabilities
5. **Collaborative features**: Teams, SharePoint integration

### Scalability Considerations
- **Connection pooling**: Already implemented via httpx
- **Token refresh optimization**: Background refresh before expiration
- **Caching strategies**: Response caching for frequently accessed data
- **Resource management**: Connection limits, timeout configuration

This implementation provides a robust, secure, and comprehensive interface to Microsoft 365 services while maintaining simplicity and reliability for AI assistant integration.
