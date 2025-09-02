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
- Additional token caching at application level for improved performance
- Works seamlessly with msgraph.GraphServiceClient

Authentication Flow:
- Uses InteractiveBrowserCredential which opens a browser for user sign-in
- Implements PKCE (Proof Key for Code Exchange) for security
- No special permissions required (unlike device flow)
- Handles token refresh automatically
- Caches tokens locally to avoid unnecessary re-authentication

Token Caching:
- Tokens are cached in ~/.microsoft_mcp_delegated_token_cache.json
- Cached tokens are validated before use (5-minute expiration buffer)
- Additional API verification ensures tokens are actually valid
- Invalid/expired tokens trigger automatic re-authentication
- Cache can be cleared manually using clear_token_cache()

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
- MICROSOFT_MCP_REDIRECT_URI environment variable (optional, for non-localhost deployments)
- Web browser available for interactive authentication
"""

import os
import asyncio
import json
import time
import logging
from pathlib import Path
from typing import NamedTuple, Optional
from dotenv import load_dotenv
from azure.identity import InteractiveBrowserCredential
from azure.core.credentials import AccessToken
from msgraph import GraphServiceClient

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

load_dotenv()

# Token cache file location
TOKEN_CACHE_FILE = Path.home() / ".microsoft_mcp_delegated_token_cache.json"

# Delegated permissions (scopes) for accessing user data on behalf of the signed-in user
# These match the scopes used in your working example.py
SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Mail.Read",
    "Team.ReadBasic.All",
    "TeamMember.ReadWrite.All",
    "Calendars.Read",
    "Files.Read",
]


class Account(NamedTuple):
    username: str
    account_id: str


class CachedToken(NamedTuple):
    """Cached access token with expiration time"""

    token: str
    expires_on: float  # Unix timestamp


def _read_token_cache() -> Optional[CachedToken]:
    """Read cached token from file"""
    try:
        if TOKEN_CACHE_FILE.exists():
            logger.info("Reading token from cache file")
            with open(TOKEN_CACHE_FILE, "r") as f:
                data = json.load(f)
                cached_token = CachedToken(
                    token=data["token"], expires_on=data["expires_on"]
                )
                logger.info(
                    f"Token cache found, expires at {time.ctime(cached_token.expires_on)}"
                )
                return cached_token
        else:
            logger.info("No token cache file found")
    except (FileNotFoundError, json.JSONDecodeError, KeyError) as e:
        logger.warning(f"Failed to read token cache: {e}")
    return None


def _write_token_cache(token: str, expires_on: float) -> None:
    """Write token to cache file"""
    try:
        TOKEN_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_CACHE_FILE, "w") as f:
            json.dump({"token": token, "expires_on": expires_on}, f)
        logger.info(f"Token cached successfully, expires at {time.ctime(expires_on)}")
    except Exception as e:
        logger.warning(f"Failed to write token cache: {e}")


def _is_token_valid(cached_token: CachedToken) -> bool:
    """Check if cached token is still valid (with 5-minute buffer)"""
    current_time = time.time()
    buffer_time = 300  # 5 minutes buffer
    is_valid = cached_token.expires_on > (current_time + buffer_time)
    logger.info(
        f"Token validity check: {'valid' if is_valid else 'expired'} (expires: {time.ctime(cached_token.expires_on)})"
    )
    return is_valid


async def _verify_token_with_api(token: str) -> bool:
    """Verify token is valid by making a simple API call"""
    try:
        logger.info("Verifying cached token with Microsoft Graph API")
        import httpx

        headers = {"Authorization": f"Bearer {token}"}
        async with httpx.AsyncClient() as client:
            response = await client.get(
                "https://graph.microsoft.com/v1.0/me", headers=headers, timeout=10.0
            )
            is_valid = response.status_code == 200
            logger.info(
                f"API token verification: {'successful' if is_valid else 'failed'} (status: {response.status_code})"
            )
            return is_valid
    except Exception as e:
        logger.warning(f"Token API verification failed: {e}")
        return False


def get_credential() -> InteractiveBrowserCredential:
    """
    Create and configure InteractiveBrowserCredential for delegated access.
    This credential handles the interactive authentication flow automatically.
    """
    logger.info("Creating InteractiveBrowserCredential for delegated access")

    client_id = os.getenv("MICROSOFT_MCP_CLIENT_ID")
    if not client_id:
        logger.error("MICROSOFT_MCP_CLIENT_ID environment variable not found")
        raise ValueError("MICROSOFT_MCP_CLIENT_ID environment variable is required")

    tenant_id = os.getenv("MICROSOFT_MCP_TENANT_ID", "common")
    redirect_uri = os.getenv("MICROSOFT_MCP_REDIRECT_URI")

    logger.info(f"Using tenant ID: {tenant_id}")
    if redirect_uri:
        logger.info(f"Using custom redirect URI: {redirect_uri}")
    else:
        logger.info("Using default localhost redirect URI")

    # Configure credential with optional redirect URI
    credential_kwargs = {
        "client_id": client_id,
        "tenant_id": tenant_id,
    }

    # Add redirect_uri if specified (for non-localhost deployments)
    if redirect_uri:
        credential_kwargs["redirect_uri"] = redirect_uri

    credential = InteractiveBrowserCredential(**credential_kwargs)
    logger.info("InteractiveBrowserCredential created successfully")

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

    client = GraphServiceClient(credentials=credential, scopes=requested_scopes)

    return client


def get_token(account_id: str | None = None) -> str:
    """
    Get an access token for Microsoft Graph API calls with caching.

    Args:
        account_id: Not used in delegated access mode, but kept for compatibility.
                   In delegated access, we always use the currently signed-in user.

    Returns:
        Valid access token for Microsoft Graph API.
    """
    logger.info("Requesting access token for Microsoft Graph API")

    # Check if we have a cached token that hasn't expired
    cached_token = _read_token_cache()
    if cached_token and _is_token_valid(cached_token):
        logger.info("Found valid cached token, verifying with API")
        # Verify the token is actually valid with the API
        try:
            import asyncio

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                is_valid = loop.run_until_complete(
                    _verify_token_with_api(cached_token.token)
                )
                if is_valid:
                    logger.info(
                        "Cached token verified successfully, using cached token"
                    )
                    return cached_token.token
                else:
                    logger.info(
                        "Cached token failed API verification, will acquire new token"
                    )
            finally:
                loop.close()
        except Exception as e:
            logger.warning(f"Token verification failed: {e}, will acquire new token")

    # No valid cached token, get a new one
    logger.info("Acquiring new access token with interactive authentication")
    credential = get_credential()

    try:
        logger.info(f"Requesting token for scopes: {', '.join(SCOPES)}")
        # Get token for specific Microsoft Graph scopes
        token: AccessToken = credential.get_token(*SCOPES)

        logger.info("Access token acquired successfully")
        # Cache the new token
        _write_token_cache(token.token, token.expires_on)

        return token.token
    except Exception as e:
        logger.error(f"Failed to acquire access token: {e}")
        # If token acquisition fails, try to clear cache and raise the error
        if TOKEN_CACHE_FILE.exists():
            logger.info("Clearing token cache due to authentication failure")
            TOKEN_CACHE_FILE.unlink()
        raise Exception(f"Failed to acquire access token: {str(e)}")


def clear_token_cache() -> None:
    """Clear the cached token to force re-authentication"""
    try:
        if TOKEN_CACHE_FILE.exists():
            TOKEN_CACHE_FILE.unlink()
            logger.info("Token cache cleared successfully")
        else:
            logger.info("No token cache file to clear")
    except Exception as e:
        logger.warning(f"Failed to clear token cache: {e}")


async def get_user_info() -> dict:
    """
    Get user information using delegated access.
    This demonstrates accessing user data on behalf of the signed-in user.

    Returns:
        Dictionary containing user information from Microsoft Graph /me endpoint.
    """
    logger.info("Retrieving user information from Microsoft Graph")
    client = get_graph_client(scopes=["User.Read"])

    try:
        me = await client.me.get()
        if me:
            user_info = {
                "displayName": me.display_name,
                "mail": me.mail or me.user_principal_name,
                "jobTitle": me.job_title,
                "id": me.id,
                "userPrincipalName": me.user_principal_name,
                "givenName": me.given_name,
                "surname": me.surname,
            }
            logger.info(
                f"User information retrieved successfully for: {user_info.get('mail', 'unknown')}"
            )
            return user_info
        else:
            logger.error("Failed to retrieve user information: empty response")
            raise Exception("Failed to retrieve user information")
    except Exception as e:
        logger.error(f"Error getting user info: {e}")
        raise Exception(f"Error getting user info: {str(e)}")


async def authenticate_new_account() -> Optional[Account]:
    """
    Authenticate a new account interactively using delegated access.
    This allows the app to act on behalf of the signed-in user.
    Uses InteractiveBrowserCredential with authorization code flow + PKCE.
    """
    logger.info("Starting new account authentication with delegated access")

    print("\nDelegated Access Authentication:")
    print("This will allow the app to access Microsoft Graph on your behalf.")
    print("Opening browser for interactive authentication...")
    print("You will be redirected to sign in with your Microsoft account.")
    print("Requested permissions:")
    for scope in SCOPES:
        print(f"   - {scope}")
    print("\nStarting authentication...")

    try:
        logger.info("Triggering interactive authentication flow")
        # Get user info to verify authentication worked
        # This will trigger the interactive authentication if needed
        user_info = await get_user_info()

        account = Account(
            username=user_info["mail"] or user_info["userPrincipalName"],
            account_id=user_info["id"],
        )

        logger.info(f"Authentication successful for account: {account.username}")
        return account
    except Exception as e:
        logger.error(f"Authentication failed: {e}")
        raise Exception(f"Authentication failed: {str(e)}")


async def list_accounts_async() -> list[Account]:
    """
    List authenticated accounts. With InteractiveBrowserCredential,
    we can only check if current authentication works.
    """
    logger.info("Listing authenticated accounts")
    try:
        user_info = await get_user_info()
        accounts = [
            Account(
                username=user_info["mail"] or user_info["userPrincipalName"],
                account_id=user_info["id"],
            )
        ]
        logger.info(f"Found {len(accounts)} authenticated account(s)")
        return accounts
    except Exception as e:
        logger.warning(f"Failed to list accounts: {e}")
        return []


def list_accounts() -> list[Account]:
    """
    Synchronous wrapper for list_accounts_async.
    """
    logger.info("Listing accounts (synchronous wrapper)")
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(list_accounts_async())
        finally:
            loop.close()
    except Exception as e:
        logger.warning(f"Failed to list accounts: {e}")
        return []
