"""
Unit tests for Microsoft Graph authentication module.
"""

import os
import json
import pytest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path
from azure.identity import AuthenticationRecord
from azure.core.credentials import AccessToken

from src.microsoft_mcp.auth import AzureAuthentication, SCOPES


class TestAzureAuthentication:
    """Test cases for AzureAuthentication class."""

    def setup_method(self):
        """Set up test fixtures before each test method."""
        self.temp_auth_file = Path("/tmp/test_auth_record.json")
        # Clean up any existing test file
        if self.temp_auth_file.exists():
            self.temp_auth_file.unlink()

    def teardown_method(self):
        """Clean up after each test method."""
        if self.temp_auth_file.exists():
            self.temp_auth_file.unlink()

    @patch.dict(os.environ, {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    def test_init_with_default_auth_file(self):
        """Test initialization with default auth record file path."""
        auth = AzureAuthentication()
        expected_path = Path.home() / ".azure-graph-auth.json"
        assert auth.auth_record_file == expected_path

    def test_init_with_custom_auth_file(self):
        """Test initialization with custom auth record file path."""
        custom_path = Path("/custom/path/auth.json")
        auth = AzureAuthentication(auth_record_file=custom_path)
        assert auth.auth_record_file == custom_path

    def test_scopes_configuration(self):
        """Test that required scopes are properly configured."""
        expected_scopes = [
            "User.Read",
            "User.ReadBasic.All",
            "Chat.Read",
            "Mail.Read",
            "Team.ReadBasic.All",
            "TeamMember.ReadWrite.All",
            "Calendars.Read",
            "Files.Read",
        ]
        assert SCOPES == expected_scopes

    @patch.dict(os.environ, {}, clear=True)
    def test_get_credential_missing_client_id(self):
        """Test that missing client ID raises ValueError."""
        auth = AzureAuthentication()
        with pytest.raises(
            ValueError, match="MICROSOFT_MCP_CLIENT_ID environment variable is required"
        ):
            auth.get_credential()

    @patch.dict(os.environ, {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_get_credential_with_minimal_config(self, mock_credential_class):
        """Test credential creation with minimal configuration."""
        mock_credential = Mock()
        mock_credential_class.return_value = mock_credential

        auth = AzureAuthentication()
        credential = auth.get_credential()

        assert credential == mock_credential
        mock_credential_class.assert_called_once()
        call_args = mock_credential_class.call_args[1]
        assert call_args["client_id"] == "test-client-id"
        # Check that tenant_id is set (could be "common" or actual tenant ID from env)
        assert "tenant_id" in call_args

    @patch.dict(
        os.environ,
        {
            "MICROSOFT_MCP_CLIENT_ID": "test-client-id",
            "MICROSOFT_MCP_TENANT_ID": "test-tenant-id",
            "MICROSOFT_MCP_REDIRECT_URI": "http://localhost:8080/callback",
        },
    )
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_get_credential_with_full_config(self, mock_credential_class):
        """Test credential creation with full configuration."""
        mock_credential = Mock()
        mock_credential_class.return_value = mock_credential

        auth = AzureAuthentication()
        credential = auth.get_credential()

        call_args = mock_credential_class.call_args[1]
        assert call_args["client_id"] == "test-client-id"
        assert call_args["tenant_id"] == "test-tenant-id"
        assert call_args["redirect_uri"] == "http://localhost:8080/callback"

    def test_read_auth_record_file_not_exists(self):
        """Test reading auth record when file doesn't exist."""
        auth = AzureAuthentication(auth_record_file=self.temp_auth_file)
        result = auth._read_auth_record()
        assert result is None

    def test_write_and_read_auth_record(self):
        """Test writing and reading auth record file."""
        # Create a mock authentication record
        mock_auth_record = Mock(spec=AuthenticationRecord)
        mock_data = {
            "authority": "https://login.microsoftonline.com/common",
            "home_account_id": "test-account-id",
            "client_id": "test-client-id",
            "username": "test@example.com",
        }
        mock_auth_record.serialize.return_value = json.dumps(mock_data)

        auth = AzureAuthentication(auth_record_file=self.temp_auth_file)

        # Write the auth record
        auth._write_auth_record(mock_auth_record)

        # Verify file was created
        assert self.temp_auth_file.exists()

        # Verify file content
        with open(self.temp_auth_file, "r") as f:
            saved_data = json.load(f)
        assert saved_data == mock_data

    def test_exists_valid_token_no_auth_record(self):
        """Test exists_valid_token when no auth record file exists."""
        auth = AzureAuthentication(auth_record_file=self.temp_auth_file)
        assert auth.exists_valid_token() is False

    @patch.dict(os.environ, {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_exists_valid_token_with_valid_token(self, mock_credential_class):
        """Test exists_valid_token when valid token is available."""
        mock_credential = Mock()
        mock_token = AccessToken("valid-token", 9999999999)  # Far future expiration
        mock_credential.get_token.return_value = mock_token
        mock_credential_class.return_value = mock_credential

        # Create a fake auth record file
        self.temp_auth_file.parent.mkdir(parents=True, exist_ok=True)
        self.temp_auth_file.write_text('{"test": "data"}')

        auth = AzureAuthentication(auth_record_file=self.temp_auth_file)
        result = auth.exists_valid_token()

        assert result is True

    @patch.dict(os.environ, {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_get_token_success(self, mock_credential_class):
        """Test successful token acquisition."""
        mock_credential = Mock()
        mock_token = AccessToken("test-access-token", 9999999999)
        mock_credential.get_token.return_value = mock_token
        mock_credential_class.return_value = mock_credential

        auth = AzureAuthentication()
        token = auth.get_token()

        assert token == "test-access-token"
        mock_credential.get_token.assert_called_with(*SCOPES)

    @patch.dict(os.environ, {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_get_token_with_details_success(self, mock_credential_class):
        """Test successful token acquisition with details."""
        mock_credential = Mock()
        expires_on = 9999999999
        mock_token = AccessToken("test-access-token", expires_on)
        mock_credential.get_token.return_value = mock_token
        mock_credential_class.return_value = mock_credential

        auth = AzureAuthentication()
        token, expiry = auth.get_token_with_details()

        assert token == "test-access-token"
        assert expiry == expires_on

    def test_clear_cache_no_file(self):
        """Test clearing cache when no auth record file exists."""
        auth = AzureAuthentication(auth_record_file=self.temp_auth_file)
        # Should not raise an exception
        auth.clear_cache()
        assert not self.temp_auth_file.exists()

    def test_clear_cache_with_file(self):
        """Test clearing cache when auth record file exists."""
        # Create a dummy auth record file
        self.temp_auth_file.parent.mkdir(parents=True, exist_ok=True)
        self.temp_auth_file.write_text('{"test": "data"}')

        auth = AzureAuthentication(auth_record_file=self.temp_auth_file)
        auth.clear_cache()

        assert not self.temp_auth_file.exists()

    def test_clear_credential_cache(self):
        """Test clearing credential cache."""
        auth = AzureAuthentication()
        auth._credential_instance = Mock()

        auth.clear_credential_cache()

        assert auth._credential_instance is None

    @patch.dict(os.environ, {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.GraphServiceClient")
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_get_graph_client(self, mock_credential_class, mock_graph_client_class):
        """Test creating Graph service client."""
        mock_credential = Mock()
        mock_credential_class.return_value = mock_credential
        mock_client = Mock()
        mock_graph_client_class.return_value = mock_client

        auth = AzureAuthentication()
        client = auth.get_graph_client()

        assert client == mock_client
        mock_graph_client_class.assert_called_once_with(
            credentials=mock_credential, scopes=SCOPES
        )

    @patch.dict(os.environ, {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.GraphServiceClient")
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_get_graph_client_custom_scopes(
        self, mock_credential_class, mock_graph_client_class
    ):
        """Test creating Graph service client with custom scopes."""
        mock_credential = Mock()
        mock_credential_class.return_value = mock_credential
        mock_client = Mock()
        mock_graph_client_class.return_value = mock_client

        custom_scopes = ["User.Read", "Mail.Read"]
        auth = AzureAuthentication()
        client = auth.get_graph_client(scopes=custom_scopes)

        mock_graph_client_class.assert_called_once_with(
            credentials=mock_credential, scopes=custom_scopes
        )
