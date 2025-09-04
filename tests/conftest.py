"""
Test configuration and fixtures for Microsoft MCP tests.
"""

import pytest
from unittest.mock import Mock, patch
import os


@pytest.fixture
def mock_auth():
    """Fixture providing a mock authentication instance."""
    auth = Mock()
    auth.exists_valid_token.return_value = True
    auth.get_token.return_value = "mock-access-token"
    return auth


@pytest.fixture
def clean_env():
    """Fixture that provides a clean environment for testing."""
    # Store original environment variables
    original_env = dict(os.environ)

    # Clear Microsoft-related environment variables for testing
    env_vars_to_clear = [
        "MICROSOFT_MCP_CLIENT_ID",
        "MICROSOFT_MCP_TENANT_ID",
        "MICROSOFT_MCP_REDIRECT_URI",
    ]

    for var in env_vars_to_clear:
        if var in os.environ:
            del os.environ[var]

    yield

    # Restore original environment
    os.environ.clear()
    os.environ.update(original_env)


@pytest.fixture
def sample_user_data():
    """Fixture providing sample user data for testing."""
    return {
        "id": "12345",
        "displayName": "John Doe",
        "mail": "john.doe@company.com",
        "userPrincipalName": "john.doe@company.com",
        "givenName": "John",
        "surname": "Doe",
        "jobTitle": "Software Engineer",
        "department": "Engineering",
        "companyName": "Test Company",
    }


@pytest.fixture
def sample_email_data():
    """Fixture providing sample email data for testing."""
    return [
        {
            "id": "email1",
            "subject": "Test Email 1",
            "from": {
                "emailAddress": {"address": "sender1@test.com", "name": "Sender One"}
            },
            "toRecipients": [{"emailAddress": {"address": "recipient@test.com"}}],
            "receivedDateTime": "2024-09-01T10:00:00Z",
            "hasAttachments": False,
            "isRead": False,
            "conversationId": "conv1",
        },
        {
            "id": "email2",
            "subject": "Test Email 2",
            "from": {
                "emailAddress": {"address": "sender2@test.com", "name": "Sender Two"}
            },
            "toRecipients": [{"emailAddress": {"address": "recipient@test.com"}}],
            "receivedDateTime": "2024-09-01T11:00:00Z",
            "hasAttachments": True,
            "isRead": True,
            "conversationId": "conv2",
        },
    ]
