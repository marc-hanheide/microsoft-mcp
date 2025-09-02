"""
Microsoft Graph Authentication Module - Delegated Access

This module implements delegated access authentication for Microsoft Graph API.
Delegated access allows the application to act on behalf of a signed-in user,
accessing only the data that the user has permission to access.

Key Features:
- Uses azure.identity.InteractiveBrowserCredential for modern authentication
- Implements interactive authentication with authorization code flow + PKCE
- Requests specific delegated permissions (scopes) rather than broad access
- Supports token caching automatically through azure.identity
- Works seamlessly with msgraph.GraphServiceClient

Authentication Flow:
- Uses InteractiveBrowserCredential which opens a browser for user sign-in
- Implements PKCE (Proof Key for Code Exchange) for security
- No special permissions required (unlike device flow)
- Handles token refresh automatically

Delegated Permissions Used:
- User.Read: Read the signed-in user's profile
- User.ReadBasic.All: Read basic info of all users
- Mail.Read: Read user's mail
- Mail.Send: Send mail as user
- Team.ReadBasic.All: Read basic team information
- TeamMember.ReadWrite.All: Read and write team membership

This is different from "app-only access" where the app acts with its own identity
and requires application permissions rather than delegated permissions.

Requirements:
- Azure AD app registration with public client flow enabled
- Delegated permissions configured in Azure AD
- MICROSOFT_MCP_CLIENT_ID and MICROSOFT_MCP_TENANT_ID environment variables
- Web browser available for interactive authentication
"""

import os
import asyncio
from typing import NamedTuple, Optional
from dotenv import load_dotenv
from azure.identity import InteractiveBrowserCredential
from msgraph import GraphServiceClient

load_dotenv()

# Delegated permissions (scopes) for accessing user data on behalf of the signed-in user
# These match the scopes used in your working example.py
SCOPES = [
    "User.Read",
    "User.ReadBasic.All", 
    "Mail.Read",
    "Mail.Send",
    "Team.ReadBasic.All",
    "TeamMember.ReadWrite.All",
    "Calendars.Read",
    "Files.Read"
]


class Account(NamedTuple):
    username: str
    account_id: str


def get_credential() -> InteractiveBrowserCredential:
    """
    Create and configure InteractiveBrowserCredential for delegated access.
    This credential handles the interactive authentication flow automatically.
    """
    client_id = os.getenv("MICROSOFT_MCP_CLIENT_ID")
    if not client_id:
        raise ValueError("MICROSOFT_MCP_CLIENT_ID environment variable is required")

    tenant_id = os.getenv("MICROSOFT_MCP_TENANT_ID", "common")
    
    credential = InteractiveBrowserCredential(
        client_id=client_id,
        tenant_id=tenant_id,
    )
    
    return credential


def get_graph_client(scopes: Optional[list[str]] = None) -> GraphServiceClient:
    """
    Get a configured Microsoft Graph client for delegated access.
    
    Args:
        scopes: Custom scopes to request. If None, uses default SCOPES.
    
    Returns:
        GraphServiceClient configured for delegated access.
    """
    credential = get_credential()
    requested_scopes = scopes or SCOPES
    
    client = GraphServiceClient(
        credentials=credential, 
        scopes=requested_scopes
    )
    
    return client


async def get_user_info() -> dict:
    """
    Get user information using delegated access.
    This demonstrates accessing user data on behalf of the signed-in user.
    
    Returns:
        Dictionary containing user information from Microsoft Graph /me endpoint.
    """
    client = get_graph_client(scopes=["User.Read"])
    
    try:
        me = await client.me.get()
        if me:
            return {
                "displayName": me.display_name,
                "mail": me.mail or me.user_principal_name,
                "jobTitle": me.job_title,
                "id": me.id,
                "userPrincipalName": me.user_principal_name,
                "givenName": me.given_name,
                "surname": me.surname
            }
        else:
            raise Exception("Failed to retrieve user information")
    except Exception as e:
        raise Exception(f"Error getting user info: {str(e)}")


async def authenticate_new_account() -> Optional[Account]:
    """
    Authenticate a new account interactively using delegated access.
    This allows the app to act on behalf of the signed-in user.
    Uses InteractiveBrowserCredential with authorization code flow + PKCE.
    """
    print("\nDelegated Access Authentication:")
    print("This will allow the app to access Microsoft Graph on your behalf.")
    print("Opening browser for interactive authentication...")
    print("You will be redirected to sign in with your Microsoft account.")
    print("Requested permissions:")
    for scope in SCOPES:
        print(f"   - {scope}")
    print("\nStarting authentication...")
    
    try:
        # Get user info to verify authentication worked
        # This will trigger the interactive authentication if needed
        user_info = await get_user_info()
        
        return Account(
            username=user_info["mail"] or user_info["userPrincipalName"],
            account_id=user_info["id"]
        )
    except Exception as e:
        raise Exception(f"Authentication failed: {str(e)}")


async def list_accounts_async() -> list[Account]:
    """
    List authenticated accounts. With InteractiveBrowserCredential, 
    we can only check if current authentication works.
    """
    try:
        user_info = await get_user_info()
        return [Account(
            username=user_info["mail"] or user_info["userPrincipalName"],
            account_id=user_info["id"]
        )]
    except:
        return []


def list_accounts() -> list[Account]:
    """
    Synchronous wrapper for list_accounts_async.
    """
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(list_accounts_async())
        finally:
            loop.close()
    except:
        return []
