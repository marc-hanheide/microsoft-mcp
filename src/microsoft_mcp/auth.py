"""
Microsoft Graph Authentication Module - Delegated Access

This module implements simplified delegated access authentication for Microsoft Graph API.
Delegated access allows the application to act on behalf of a signed-in user,
accessing only the data that the user has permission to access.

Key Features:
- Uses azure.identity.InteractiveBrowserCredential for modern authentication
- Leverages Azure SDK's built-in token caching and refresh token handling
- Uses AuthenticationRecord for persistent authentication across sessions
- No manual token management or background refresh services needed
- Works seamlessly with msgraph.GraphServiceClient

Authentication Flow:
- Uses InteractiveBrowserCredential with persistent token cache
- First authentication saves an AuthenticationRecord to ~/.azure-graph-auth.json
- Subsequent runs use the saved AuthenticationRecord for silent authentication
- Azure SDK handles all token refresh automatically

Token Caching:
- Tokens are cached by Azure SDK using platform-specific secure storage
- AuthenticationRecord enables silent authentication across application restarts
- No manual token validation or refresh needed

Delegated Permissions Used:
- User.Read: Read the signed-in user's profile
- User.ReadBasic.All: Read basic info of all users
- Chat.Read: Read user's chat messages
- Mail.Read: Read user's mail
- Team.ReadBasic.All: Read basic team information
- TeamMember.ReadWrite.All: Read and write team membership
- Calendars.Read: Read user's calendar
- Files.Read: Read user's files

Requirements:
- Azure AD app registration with public client flow enabled
- Delegated permissions configured in Azure AD
- MICROSOFT_MCP_CLIENT_ID and MICROSOFT_MCP_TENANT_ID environment variables
- MICROSOFT_MCP_REDIRECT_URI environment variable (optional, for non-localhost deployments)
- Web browser available for interactive authentication
"""

import os
import json
import logging
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv
from azure.identity import (
    InteractiveBrowserCredential,
    TokenCachePersistenceOptions,
    AuthenticationRecord,
)
from azure.core.credentials import AccessToken
from msgraph import GraphServiceClient

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

load_dotenv()

# Delegated permissions (scopes) for accessing user data on behalf of the signed-in user
SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Chat.Read",
    "Mail.Read",
    "Team.ReadBasic.All",
    "TeamMember.ReadWrite.All",
    "Calendars.Read",
    "Files.Read",
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


class AzureAuthentication:
    """
    Simplified Azure Authentication class for Microsoft Graph API with delegated access.

    This class leverages Azure SDK's built-in token caching and AuthenticationRecord
    functionality to provide seamless authentication across application restarts.
    """

    def __init__(
        self,
        auth_record_file: Optional[Path] = None,
    ):
        """
        Initialize the Azure Authentication instance.

        Args:
            auth_record_file: Path to AuthenticationRecord file (defaults to ~/.azure-graph-auth.json)
        """
        self.auth_record_file = auth_record_file or (
            Path.home() / ".azure-graph-auth.json"
        )
        self._credential_instance = None

    def _read_auth_record(self) -> Optional[AuthenticationRecord]:
        """Read AuthenticationRecord from file"""
        try:
            if self.auth_record_file.exists():
                logger.info("Reading AuthenticationRecord from file")
                with open(self.auth_record_file, "r") as f:
                    auth_record_data = json.load(f)
                    auth_record = AuthenticationRecord.deserialize(
                        json.dumps(auth_record_data)
                    )
                    logger.info("AuthenticationRecord loaded successfully")
                    return auth_record
            else:
                logger.info("No AuthenticationRecord file found")
        except Exception as e:
            logger.warning(f"Failed to read AuthenticationRecord: {e}")
        return None

    def _write_auth_record(self, auth_record: AuthenticationRecord) -> None:
        """Write AuthenticationRecord to file"""
        try:
            self.auth_record_file.parent.mkdir(parents=True, exist_ok=True)
            auth_record_data = json.loads(auth_record.serialize())

            with open(self.auth_record_file, "w") as f:
                json.dump(auth_record_data, f, indent=2)
            logger.info("AuthenticationRecord saved successfully")
        except Exception as e:
            logger.warning(f"Failed to write AuthenticationRecord: {e}")

    def get_credential(self) -> InteractiveBrowserCredential:
        """
        Create and configure InteractiveBrowserCredential for delegated access.
        Uses persistent token cache and AuthenticationRecord for seamless re-authentication.
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

        # Configure persistent token cache
        token_cache = TokenCachePersistenceOptions(
            allow_unencrypted_storage=True, name="/Users/mhanheide/.azure-token-cache"
        )

        # Try to load existing AuthenticationRecord
        auth_record = self._read_auth_record()

        credential_kwargs = {
            "client_id": client_id,
            "tenant_id": tenant_id,
            "cache_persistence_options": token_cache,
        }

        # Add existing authentication record if available
        if auth_record:
            credential_kwargs["authentication_record"] = auth_record
            logger.info("Using existing AuthenticationRecord for silent authentication")

        # Add redirect_uri if specified (for non-localhost deployments)
        if redirect_uri:
            credential_kwargs["redirect_uri"] = redirect_uri

        self._credential_instance = InteractiveBrowserCredential(**credential_kwargs)
        logger.info("InteractiveBrowserCredential created successfully")

        return self._credential_instance

    def authenticate(self) -> AuthenticationRecord:
        """
        Perform interactive authentication and save AuthenticationRecord for future use.
        This method should be called at least once to establish the authentication record.
        """
        logger.info("Performing interactive authentication")
        credential = self.get_credential()

        # Authenticate and get the AuthenticationRecord
        auth_record = credential.authenticate(scopes=SCOPES)

        # Save the AuthenticationRecord for future use
        self._write_auth_record(auth_record)

        logger.info("Authentication completed and record saved")
        return auth_record

    def get_token_with_details(self) -> tuple[str, int]:
        """
        Get an access token along with its expiration timestamp.
        Uses Azure SDK's built-in caching and refresh token handling.

        Returns:
            Tuple of (token_string, expires_on_timestamp)
        """
        logger.info("Requesting access token with details for Microsoft Graph API")

        credential = self.get_credential()

        try:
            logger.info(f"Requesting token for scopes: {', '.join(SCOPES)}")
            token: AccessToken = credential.get_token(*SCOPES)

            logger.info("Access token acquired successfully")
            return token.token, token.expires_on

        except Exception as e:
            logger.error(f"Failed to acquire access token: {e}")
            # If authentication fails and we don't have an auth record, try interactive auth
            if not self.auth_record_file.exists():
                logger.info(
                    "No AuthenticationRecord found, attempting interactive authentication"
                )
                try:
                    self.authenticate()
                    # Retry token acquisition after authentication
                    token: AccessToken = credential.get_token(*SCOPES)
                    logger.info(
                        "Access token acquired after interactive authentication"
                    )
                    return token.token, token.expires_on
                except Exception as retry_e:
                    logger.error(
                        f"Failed to acquire token after interactive authentication: {retry_e}"
                    )
                    raise
            else:
                # Auth record exists but token acquisition failed, clear cache and retry
                logger.info("Clearing cached data and retrying authentication")
                self.clear_cache()
                self._credential_instance = None
                raise Exception(f"Failed to acquire access token: {str(e)}")

    def get_token(self) -> str:
        """
        Get an access token for Microsoft Graph API calls.
        Uses Azure SDK's built-in caching and refresh token handling.

        Returns:
            Valid access token for Microsoft Graph API.
        """
        logger.info("Requesting access token for Microsoft Graph API")

        credential = self.get_credential()

        try:
            logger.info(f"Requesting token for scopes: {', '.join(SCOPES)}")
            token: AccessToken = credential.get_token(*SCOPES)

            logger.info("Access token acquired successfully")
            return token.token

        except Exception as e:
            logger.error(f"Failed to acquire access token: {e}")
            # If authentication fails and we don't have an auth record, try interactive auth
            if not self.auth_record_file.exists():
                logger.info(
                    "No AuthenticationRecord found, attempting interactive authentication"
                )
                try:
                    self.authenticate()
                    # Retry token acquisition after authentication
                    token: AccessToken = credential.get_token(*SCOPES)
                    logger.info(
                        "Access token acquired after interactive authentication"
                    )
                    return token.token
                except Exception as retry_e:
                    logger.error(
                        f"Failed to acquire token after interactive authentication: {retry_e}"
                    )
                    raise
            else:
                # Auth record exists but token acquisition failed, clear cache and retry
                logger.info("Clearing cached data and retrying authentication")
                self.clear_cache()
                self._credential_instance = None
                raise Exception(f"Failed to acquire access token: {str(e)}")

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
        Check if a valid access token can be obtained silently.
        This doesn't guarantee the token is valid, but indicates if silent auth is possible.
        """
        try:
            if not self.auth_record_file.exists():
                return False

            credential = self.get_credential()
            # Try to get a token silently
            token: AccessToken = credential.get_token(*SCOPES)
            return token is not None
        except Exception:
            return False

    def clear_cache(self) -> None:
        """Clear the AuthenticationRecord and force re-authentication"""
        try:
            if self.auth_record_file.exists():
                self.auth_record_file.unlink()
                logger.info("AuthenticationRecord cleared successfully")
            else:
                logger.info("No AuthenticationRecord file to clear")

            # Clear the credential instance to force new authentication
            self._credential_instance = None
            logger.info("Credential instance cleared")

        except Exception as e:
            logger.warning(f"Failed to clear cache: {e}")

    def clear_credential_cache(self) -> None:
        """Clear the credential instance to force re-authentication"""
        self._credential_instance = None
        logger.info("Credential instance cleared")
