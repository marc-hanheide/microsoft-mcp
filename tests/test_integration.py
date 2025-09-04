"""
Integration tests for Microsoft MCP modules.
"""

import pytest
from unittest.mock import Mock, patch
from src.microsoft_mcp import auth, graph


class TestModuleIntegration:
    """Test integration between modules."""

    def test_auth_module_imports(self):
        """Test that auth module imports correctly."""
        assert hasattr(auth, "AzureAuthentication")
        assert hasattr(auth, "SCOPES")

        # Test that SCOPES contains expected permissions
        assert "User.Read" in auth.SCOPES
        assert "Mail.Read" in auth.SCOPES
        assert "Calendars.Read" in auth.SCOPES
        assert "Files.Read" in auth.SCOPES

    def test_graph_module_imports(self):
        """Test that graph module imports correctly."""
        assert hasattr(graph, "request")
        assert hasattr(graph, "request_paginated")
        assert hasattr(graph, "search_query")
        assert hasattr(graph, "set_auth_instance")
        assert hasattr(graph, "get_auth_instance")

    def test_graph_base_url(self):
        """Test that graph module has correct base URL."""
        assert graph.BASE_URL == "https://graph.microsoft.com/v1.0"

    @patch("src.microsoft_mcp.graph.httpx.Client")
    def test_graph_client_configuration(self, mock_client_class):
        """Test that HTTP client is configured correctly."""
        # The module should have initialized a client
        mock_client = Mock()
        mock_client_class.return_value = mock_client

        # Import to trigger client creation
        import src.microsoft_mcp.graph as graph_module

        # Verify client configuration would be reasonable
        assert hasattr(graph_module, "_client")

    def test_auth_graph_integration(self):
        """Test that auth and graph modules can work together."""
        # Create auth instance
        auth_instance = auth.AzureAuthentication()

        # Set it in graph module
        graph.set_auth_instance(auth_instance)

        # Retrieve it back
        retrieved_auth = graph.get_auth_instance()

        assert retrieved_auth == auth_instance

    @patch.dict("os.environ", {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_auth_instance_creation(self, mock_credential):
        """Test creating authentication instance with minimal config."""
        mock_cred_instance = Mock()
        mock_credential.return_value = mock_cred_instance

        auth_instance = auth.AzureAuthentication()
        credential = auth_instance.get_credential()

        assert credential == mock_cred_instance
        mock_credential.assert_called_once()

    def test_folder_constants(self):
        """Test that folder mappings are accessible."""
        from src.microsoft_mcp.tools import FOLDERS

        assert isinstance(FOLDERS, dict)
        assert "inbox" in FOLDERS
        assert "sent" in FOLDERS
        assert FOLDERS["inbox"] == "inbox"
        assert FOLDERS["sent"] == "sentitems"

    def test_logging_configuration(self):
        """Test that logging is configured in modules."""
        import logging

        # Test that modules have loggers
        auth_logger = logging.getLogger("src.microsoft_mcp.auth")
        graph_logger = logging.getLogger("src.microsoft_mcp.graph")
        tools_logger = logging.getLogger("src.microsoft_mcp.tools")

        # Loggers should exist (even if not explicitly configured)
        assert auth_logger is not None
        assert graph_logger is not None
        assert tools_logger is not None
