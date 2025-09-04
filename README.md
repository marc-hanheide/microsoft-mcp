# Microsoft MCP

Powerful MCP server for Microsoft Graph API - a complete AI assistant toolkit for Outlook, Calendar, OneDrive, and Contacts with modern Azure SDK authentication.

## Features

- **🔐 Modern Authentication**: Azure SDK-based delegated access with automatic token management
- **📧 Email Management**: Read, send, reply, manage attachments, organize folders, date filtering
- **📅 Calendar Intelligence**: Create, update, check availability, respond to invitations
- **📁 OneDrive Files**: Upload, download, browse with pagination, large file support
- **👥 Contacts**: Search and list contacts from your address book
- **🔍 Unified Search**: Search across emails, files, events, and people
- **🗂️ Flexible Storage**: Configurable credential and token cache locations
- **🛡️ Secure Token Management**: Platform-specific secure storage via Azure SDK

## Quick Start with Claude Desktop

```bash
# Add Microsoft MCP server (replace with your Azure app ID)
claude mcp add microsoft-mcp -e MICROSOFT_MCP_CLIENT_ID=your-app-id-here -- uvx --from git+https://github.com/marc-hanheide/microsoft-mcp.git microsoft-mcp

# Start Claude Desktop
claude
```

### Usage Examples

```bash
# Email examples
> read my latest emails with full content
> reply to the email from John saying "I'll review this today"
> send an email with attachment to alice@example.com
> show emails from last week

# Calendar examples  
> show my calendar for next week
> check if I'm free tomorrow at 2pm
> create a meeting with Bob next Monday at 10am

# File examples
> list files in my OneDrive
> upload this report to OneDrive
> search for "project proposal" across all my files

# Search examples
> search for "quarterly report" across all services
> find contacts named "Smith"
```

## Available Tools

### Email Tools (9 tools)
- **`list_emails`** - List emails with optional body content and date filtering
- **`get_email`** - Get specific email with attachments
- **`create_email_draft`** - Create email draft with attachments support
- **`send_email`** - Send email immediately with CC/BCC and attachments
- **`reply_to_email`** - Reply maintaining thread context
- **`reply_all_email`** - Reply to all recipients in thread
- **`update_email`** - Mark emails as read/unread
- **`move_email`** - Move emails between folders
- **`delete_email`** - Delete emails
- **`get_attachment`** - Get email attachment content
- **`search_emails`** - Search emails by query

### Calendar Tools (7 tools)
- **`list_events`** - List calendar events with details
- **`get_event`** - Get specific event details
- **`create_event`** - Create events with location and attendees
- **`update_event`** - Reschedule or modify events
- **`delete_event`** - Cancel events
- **`respond_event`** - Accept/decline/tentative response to invitations
- **`check_availability`** - Check free/busy times for scheduling
- **`search_events`** - Search calendar events

### Contact Tools (6 tools)
- **`list_contacts`** - List all contacts
- **`get_contact`** - Get specific contact details
- **`create_contact`** - Create new contact
- **`update_contact`** - Update contact information
- **`delete_contact`** - Delete contact
- **`search_contacts`** - Search contacts by query

### File Tools (6 tools)
- **`list_files`** - Browse OneDrive files and folders
- **`get_file`** - Download file content
- **`create_file`** - Upload files to OneDrive
- **`update_file`** - Update existing file content
- **`delete_file`** - Delete files or folders
- **`search_files`** - Search files in OneDrive

### Utility Tools (2 tools)
- **`unified_search`** - Search across emails, events, and files
- **`get_user_details`** - Get current user information

## Manual Setup

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → App registrations
2. New registration → Name: `microsoft-mcp`
3. Supported account types: Personal + Work/School
4. Authentication → Allow public client flows: Yes
5. API permissions → Add these delegated permissions:
   - User.Read
   - User.ReadBasic.All
   - Chat.Read
   - Mail.Read
   - Team.ReadBasic.All
   - TeamMember.ReadWrite.All
   - Calendars.Read
   - Files.Read
6. Copy Application ID

### 2. Installation

```bash
git clone https://github.com/marc-hanheide/microsoft-mcp.git
cd microsoft-mcp
uv sync
```

### 3. Authentication

#### Basic Authentication
```bash
# Set your Azure app ID
export MICROSOFT_MCP_CLIENT_ID="your-app-id-here"

# Optional: Set custom redirect URI for non-localhost deployments
# export MICROSOFT_MCP_REDIRECT_URI="https://your-app.azurewebsites.net/auth/callback"

# Run authentication script
uv run authenticate.py
```

#### Custom Credential Storage
You can specify custom locations for storing authentication credentials and tokens:

```bash
# Store credentials and tokens in custom locations
AZURE_CRED_CACHE_FILE=./creds/azure-credentials.json \
AZURE_TOKEN_CACHE_FILE=./creds/azure-token \
MICROSOFT_MCP_CLIENT_ID="your-app-id-here" \
./authenticate.py
```

**Environment Variables for Custom Storage:**
- `AZURE_CRED_CACHE_FILE`: Path to store AuthenticationRecord (authentication metadata)
- `AZURE_TOKEN_CACHE_FILE`: Base path for Azure SDK token cache (platform-specific secure storage)

This allows you to:
- Store credentials in project-specific directories
- Use different credentials for different projects
- Keep authentication data organized
- Facilitate team sharing of configuration (credentials only, not tokens)

### 4. Claude Desktop Configuration

Add to your Claude Desktop configuration:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`  
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`  
**Linux**: `~/.config/claude/claude_desktop_config.json`

#### Basic Configuration
```json
{
  "mcpServers": {
    "microsoft": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/marc-hanheide/microsoft-mcp.git", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_CLIENT_ID": "your-app-id-here"
      }
    }
  }
}
```

#### Configuration with Custom Storage
```json
{
  "mcpServers": {
    "microsoft": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/marc-hanheide/microsoft-mcp.git", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_CLIENT_ID": "your-app-id-here",
        "AZURE_CRED_CACHE_FILE": "/path/to/creds/azure-credentials.json",
        "AZURE_TOKEN_CACHE_FILE": "/path/to/creds/azure-token",
        "MICROSOFT_MCP_REDIRECT_URI": "https://your-app.azurewebsites.net/auth/callback"
      }
    }
  }
}
```

#### Local Development Configuration
```json
{
  "mcpServers": {
    "microsoft": {
      "command": "uv",
      "args": ["--directory", "/path/to/microsoft-mcp", "run", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_CLIENT_ID": "your-app-id-here",
        "AZURE_CRED_CACHE_FILE": "./creds/azure-credentials.json",
        "AZURE_TOKEN_CACHE_FILE": "./creds/azure-token"
      }
    }
  }
}
```

## Authentication & Credential Management

Microsoft MCP uses modern Azure SDK authentication with flexible credential storage options.

### Authentication Flow
1. **Interactive Browser Authentication**: Opens browser for Microsoft sign-in
2. **AuthenticationRecord Storage**: Saves authentication metadata (not tokens) for future use
3. **Azure SDK Token Management**: Handles token refresh automatically using platform-specific secure storage
4. **Silent Authentication**: Subsequent runs authenticate silently using saved AuthenticationRecord

### Storage Options

#### Default Storage Locations
- **AuthenticationRecord**: `~/.ms-graph-mcp-azure-auth-record.json`
- **Token Cache**: `~/.ms-graph-mcp-azure-token-cache.nocache` (platform-specific secure storage)

#### Custom Storage Locations
```bash
# Specify custom paths for credentials and tokens
AZURE_CRED_CACHE_FILE=./creds/azure-credentials.json \
AZURE_TOKEN_CACHE_FILE=./creds/azure-token \
./authenticate.py
```

### Token Security
- **No Sensitive Data in AuthenticationRecord**: Only contains metadata for silent authentication
- **Secure Token Storage**: Azure SDK uses platform-specific secure storage (Windows Data Protection API, macOS Keychain, etc.)
- **Automatic Refresh**: Tokens are refreshed automatically by Azure SDK
- **Manual Cache Clearing**: Use `auth.clear_cache()` to force re-authentication

### Multi-Environment Setup
You can maintain separate credentials for different environments:

```bash
# Development environment
AZURE_CRED_CACHE_FILE=./dev-creds/azure-credentials.json ./authenticate.py

# Production environment  
AZURE_CRED_CACHE_FILE=./prod-creds/azure-credentials.json ./authenticate.py
```

## Development

```bash
# Run tests
uv run pytest tests/ -v

# Type checking
uv run pyright

# Format code
uvx ruff format .

# Lint
uvx ruff check --fix --unsafe-fixes .
```

## Example: AI Assistant Scenarios

### Smart Email Management
```python
# List latest emails with full content
emails = list_emails(limit=10, include_body=True)

# List emails from specific date range
recent_emails = list_emails(
    limit=20, 
    include_body=True,
    start_date="2024-01-01T00:00:00Z",
    end_date="2024-01-31T23:59:59Z"
)

# Reply maintaining thread
reply_to_email(email_id, "Thanks for your message. I'll review and get back to you.")

# Forward with attachments
email = get_email(email_id)
attachments = [get_attachment(email_id, att["id"], "temp_file.pdf") for att in email["attachments"]]
send_email("boss@company.com", f"FW: {email['subject']}", email["body"]["content"], attachments=attachments)
```

### Intelligent Scheduling
```python
# Check availability before scheduling
availability = check_availability("2024-01-15T10:00:00Z", "2024-01-15T18:00:00Z", ["colleague@company.com"])

# Create meeting with details
create_event(
    "Project Review",
    "2024-01-15T14:00:00Z", 
    "2024-01-15T15:00:00Z",
    location="Conference Room A",
    body="Quarterly review of project progress",
    attendees=["colleague@company.com", "manager@company.com"]
)
```

### File Management
```python
# Upload and organize files
create_file("reports/quarterly_report.pdf", file_content, "application/pdf")

# Search across all services
results = unified_search("quarterly report")
```

## Security Notes

- **AuthenticationRecord**: Contains only authentication metadata, no sensitive tokens
- **Secure Token Storage**: Azure SDK manages tokens using platform-specific secure storage (Windows Data Protection API, macOS Keychain, Linux Secret Service)
- **Automatic Token Refresh**: Azure SDK handles token refresh transparently
- **Configurable Storage**: Credentials and tokens can be stored in custom locations
- **No Environment Tokens**: No tokens stored in environment variables or code
- **Delegated Permissions**: Uses delegated access (user-scoped) rather than application permissions
- **Principle of Least Privilege**: Only requests necessary permissions

## Troubleshooting

### Authentication Issues
- **Authentication fails**: Check your `MICROSOFT_MCP_CLIENT_ID` is correct
- **"Need admin approval"**: Use `MICROSOFT_MCP_TENANT_ID=consumers` for personal accounts  
- **Token errors**: Clear cache and re-authenticate:
  ```bash
  # Remove AuthenticationRecord and force re-authentication
  rm ~/.ms-graph-mcp-azure-auth-record.json
  # Or if using custom location:
  rm ./creds/azure-credentials.json
  ```
- **Permission errors**: Ensure all required API permissions are granted in Azure Portal

### Storage Issues
- **Custom storage not working**: Ensure directories exist and are writable:
  ```bash
  mkdir -p ./creds
  chmod 755 ./creds
  ```
- **Token cache issues**: Azure SDK manages token cache automatically, but you can force refresh by clearing AuthenticationRecord

### Connection Issues
- **Network timeouts**: Check internet connection and firewall settings
- **Rate limiting**: Tool automatically retries with exponential backoff
- **Browser not opening**: Ensure default browser is set and accessible

## License

MIT