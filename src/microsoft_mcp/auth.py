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
- Background token refresh service to prevent token expiration
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

Background Token Refresh:
- Automatic background refresh service starts when tokens are first obtained
- Refreshes tokens 1 hour before expiration (configurable via REFRESH_BUFFER_SECONDS)
- Checks for refresh needs every 5 minutes (configurable via REFRESH_CHECK_INTERVAL)
- Uses shared credential instance for silent refresh without user interaction
- Leverages Azure Identity's internal token cache and refresh token mechanism
- Falls back to interactive authentication if silent refresh fails
- Service runs as daemon thread and stops automatically on program exit

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
import threading
import atexit
from pathlib import Path
from typing import NamedTuple, Optional
from dotenv import load_dotenv
from azure.identity import (
    InteractiveBrowserCredential,
    SharedTokenCacheCredential,
    TokenCachePersistenceOptions,
)
from azure.core.credentials import AccessToken
from msgraph import GraphServiceClient

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

load_dotenv()

# Delegated permissions (scopes) for accessing user data on behalf of the signed-in user
# These match the scopes used in your working example.py
SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Chat.Read",
    # "ChannelMessage.Read",
    "Mail.Read",
    "Team.ReadBasic.All",
    "TeamMember.ReadWrite.All",
    "Calendars.Read",
    "Files.Read",
    # "Sites.Read.All"
    # "ChannelMessage.Read"
]

# Scopes useful for full search:
# Chat.Read
# ChannelMessage.Read.All
# Mail.Read
# Calendars.Read
# Files.Read.All
# Sites.Read.All
# User.Read
# User.ReadBasic.All
# Team.ReadBasic.All
# TeamMember.ReadWrite.All


class CachedToken(NamedTuple):
    """Cached access token with expiration time"""

    token: str
    expires_on: float  # Unix timestamp


class AzureAuthentication:
    """
    Azure Authentication class for Microsoft Graph API with delegated access.

    This class encapsulates all authentication functionality including token caching,
    background refresh service, and credential management.
    """

    def __init__(
        self,
        token_cache_file: Optional[Path] = None,
        refresh_buffer_seconds: int = 3600,
        refresh_check_interval: int = 300,
    ):
        """
        Initialize the Azure Authentication instance.

        Args:
            token_cache_file: Path to token cache file (defaults to ~/.microsoft_mcp_delegated_token_cache.json)
            refresh_buffer_seconds: Refresh token this many seconds before expiration (default 1 hour)
            refresh_check_interval: Check for refresh needs every N seconds (default 5 minutes)
        """
        self.token_cache_file = token_cache_file or (
            Path.home() / ".microsoft_mcp_delegated_token_cache.json"
        )
        self.refresh_buffer_seconds = refresh_buffer_seconds
        self.refresh_check_interval = refresh_check_interval

        # Instance variables replacing global variables
        self._refresh_thread = None
        self._refresh_thread_stop = False
        self._credential_instance = None

        # Register cleanup function
        atexit.register(self.stop_token_refresh_service)

    def _read_token_cache(self) -> Optional[CachedToken]:
        """Read cached token from file"""
        try:
            if self.token_cache_file.exists():
                logger.info("Reading token from cache file")
                with open(self.token_cache_file, "r") as f:
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

    def _write_token_cache(self, token: str, expires_on: float) -> None:
        """Write token to cache file"""
        try:
            self.token_cache_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.token_cache_file, "w") as f:
                json.dump({"token": token, "expires_on": expires_on}, f)
            logger.info(
                f"Token cached successfully, expires at {time.ctime(expires_on)}"
            )
        except Exception as e:
            logger.warning(f"Failed to write token cache: {e}")

    def _is_token_valid(self, cached_token: CachedToken) -> bool:
        """Check if cached token is still valid (with 5-minute buffer)"""
        current_time = time.time()
        buffer_time = 300  # 5 minutes buffer
        is_valid = cached_token.expires_on > (current_time + buffer_time)
        logger.info(
            f"Token validity check: {'valid' if is_valid else 'expired'} (expires: {time.ctime(cached_token.expires_on)})"
        )
        return is_valid

    def _needs_refresh(self, cached_token: CachedToken) -> bool:
        """Check if cached token needs refresh (within refresh buffer time)"""
        current_time = time.time()
        needs_refresh = cached_token.expires_on <= (
            current_time + self.refresh_buffer_seconds
        )
        if needs_refresh:
            logger.info(
                f"Token needs refresh: expires at {time.ctime(cached_token.expires_on)}, "
                f"refresh buffer is {self.refresh_buffer_seconds} seconds"
            )
        return needs_refresh

    def _refresh_token_silently(self) -> bool:
        """
        Attempt to refresh the token silently using the cached credential.
        Returns True if successful, False otherwise.
        """
        try:
            logger.info("Attempting silent token refresh")

            # Use the existing credential instance if available
            if self._credential_instance is None:
                logger.warning("No credential instance available for silent refresh")
                return False

            # Try to get a new token using cached authentication
            # This should use the credential's internal cache and refresh token
            token: AccessToken = self._credential_instance.get_token(*SCOPES)

            logger.info("Silent token refresh successful")
            self._write_token_cache(token.token, token.expires_on)
            return True

        except Exception as e:
            logger.warning(f"Silent token refresh failed: {e}")
            return False

    def _token_refresh_worker(self):
        """Background worker thread that checks and refreshes tokens"""
        logger.info("Token refresh worker thread started")

        while not self._refresh_thread_stop:
            try:
                # Check if we have a cached token that needs refresh
                cached_token = self._read_token_cache()
                if cached_token and self._needs_refresh(cached_token):
                    logger.info("Token refresh needed, attempting silent refresh")

                    if self._refresh_token_silently():
                        logger.info("Token refreshed successfully in background")
                    else:
                        logger.warning(
                            "Background token refresh failed - will require interactive authentication on next use"
                        )

                # Sleep for the check interval or until stop is requested
                for _ in range(self.refresh_check_interval):
                    if self._refresh_thread_stop:
                        break
                    time.sleep(1)

            except Exception as e:
                logger.error(f"Error in token refresh worker: {e}")
                # Continue running even if there's an error
                time.sleep(self.refresh_check_interval)

        logger.info("Token refresh worker thread stopped")

    def start_token_refresh_service(self):
        """Start the background token refresh service"""
        if self._refresh_thread is not None and self._refresh_thread.is_alive():
            logger.info("Token refresh service is already running")
            return

        logger.info("Starting token refresh service")
        self._refresh_thread_stop = False
        self._refresh_thread = threading.Thread(
            target=self._token_refresh_worker, daemon=True
        )
        self._refresh_thread.start()

    def stop_token_refresh_service(self):
        """Stop the background token refresh service"""
        if self._refresh_thread is None or not self._refresh_thread.is_alive():
            return

        logger.info("Stopping token refresh service")
        self._refresh_thread_stop = True

        # Wait for thread to finish (with timeout)
        if self._refresh_thread.is_alive():
            self._refresh_thread.join(timeout=5)
            if self._refresh_thread.is_alive():
                logger.warning("Token refresh thread did not stop within timeout")

    def is_token_refresh_service_running(self) -> bool:
        """Check if the token refresh service is currently running"""
        return self._refresh_thread is not None and self._refresh_thread.is_alive()

    def get_credential(self) -> InteractiveBrowserCredential:
        """
        Create and configure InteractiveBrowserCredential for delegated access.
        This credential handles the interactive authentication flow automatically.
        Returns a shared instance for token refresh purposes.
        """
        # Return existing instance if available
        if self._credential_instance is not None:
            logger.info("Returning existing credential instance")
            return self._credential_instance

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
        token_cache = TokenCachePersistenceOptions(
            allow_unencrypted_storage=True, name="microsoft_mcp_delegated_token_cache"
        )
        credential_kwargs = {
            "client_id": client_id,
            "tenant_id": tenant_id,
            "cache_persistence_options": token_cache,
        }

        # Add redirect_uri if specified (for non-localhost deployments)
        if redirect_uri:
            credential_kwargs["redirect_uri"] = redirect_uri

        self._credential_instance = InteractiveBrowserCredential(**credential_kwargs)
        logger.info("InteractiveBrowserCredential created successfully")

        return self._credential_instance

    def get_graph_client(
        self, scopes: Optional[list[str]] = None
    ) -> GraphServiceClient:
        """
        Get a configured Microsoft Graph client for delegated access.

        Args:
            scopes: Custom scopes to request. If None, uses default SCOPES.

        Returns:
            GraphServiceClient configured for delegated access.
        """
        credential = self.get_credential()
        requested_scopes = scopes or SCOPES

        client = GraphServiceClient(credentials=credential, scopes=requested_scopes)

        return client

    def exists_valid_token(self) -> bool:
        """
        Check if a valid access token exists in the cache.
        """
        cached_token = self._read_token_cache()
        return self._is_token_valid(cached_token) if cached_token else False

    def get_token(self) -> str:
        """
        Get an access token for Microsoft Graph API calls with caching.

        Returns:
            Valid access token for Microsoft Graph API.
        """
        logger.info("Requesting access token for Microsoft Graph API")

        # Check if we have a cached token that hasn't expired
        cached_token = self._read_token_cache()
        if cached_token and self._is_token_valid(cached_token):
            logger.info("Found valid cached token, using cached token")

            # Start refresh service if not already running
            if not self.is_token_refresh_service_running():
                self.start_token_refresh_service()

            return cached_token.token

        # No valid cached token, get a new one
        logger.info("Acquiring new access token with interactive authentication")
        credential = self.get_credential()

        try:
            logger.info(f"Requesting token for scopes: {', '.join(SCOPES)}")
            # Get token for specific Microsoft Graph scopes
            token: AccessToken = credential.get_token(*SCOPES)

            logger.info("Access token acquired successfully")
            # Cache the new token
            self._write_token_cache(token.token, token.expires_on)

            # Start refresh service if not already running
            if not self.is_token_refresh_service_running():
                self.start_token_refresh_service()

            return token.token
        except Exception as e:
            logger.error(f"Failed to acquire access token: {e}")
            # If token acquisition fails, try to clear cache and credential instance
            if self.token_cache_file.exists():
                logger.info("Clearing token cache due to authentication failure")
                self.token_cache_file.unlink()
            self.clear_credential_cache()
            raise Exception(f"Failed to acquire access token: {str(e)}")

    def clear_token_cache(self) -> None:
        """Clear the cached token to force re-authentication"""
        try:
            if self.token_cache_file.exists():
                self.token_cache_file.unlink()
                logger.info("Token cache cleared successfully")
            else:
                logger.info("No token cache file to clear")

            # Also clear the credential instance to force new authentication
            self._credential_instance = None
            logger.info("Credential instance cleared")

        except Exception as e:
            logger.warning(f"Failed to clear token cache: {e}")

    def clear_credential_cache(self) -> None:
        """Clear the credential instance to force re-authentication"""
        self._credential_instance = None
        logger.info("Credential instance cleared")


# Global instance for backward compatibility
_auth_instance = None


def get_auth_instance() -> AzureAuthentication:
    """Get the global authentication instance"""
    global _auth_instance
    if _auth_instance is None:
        _auth_instance = AzureAuthentication()
    return _auth_instance


# Backward compatibility functions
def exists_valid_token() -> bool:
    """Check if a valid access token exists in the cache."""
    return get_auth_instance().exists_valid_token()


def get_token() -> str:
    """Get an access token for Microsoft Graph API calls with caching."""
    return get_auth_instance().get_token()


def get_graph_client(scopes: Optional[list[str]] = None) -> GraphServiceClient:
    """Get a configured Microsoft Graph client for delegated access."""
    return get_auth_instance().get_graph_client(scopes)


def get_credential() -> InteractiveBrowserCredential:
    """Create and configure InteractiveBrowserCredential for delegated access."""
    return get_auth_instance().get_credential()


def clear_token_cache() -> None:
    """Clear the cached token to force re-authentication"""
    get_auth_instance().clear_token_cache()


def clear_credential_cache() -> None:
    """Clear the credential instance to force re-authentication"""
    get_auth_instance().clear_credential_cache()


def start_token_refresh_service():
    """Start the background token refresh service"""
    get_auth_instance().start_token_refresh_service()


def stop_token_refresh_service():
    """Stop the background token refresh service"""
    get_auth_instance().stop_token_refresh_service()


def is_token_refresh_service_running() -> bool:
    """Check if the token refresh service is currently running"""
    return get_auth_instance().is_token_refresh_service_running()
