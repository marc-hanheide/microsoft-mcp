"""
Unit tests for Microsoft Graph API module.
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import httpx

from src.microsoft_mcp.graph import (
    request,
    request_paginated,
    search_query,
    set_auth_instance,
    get_auth_instance,
    BASE_URL,
)
from src.microsoft_mcp.auth import AzureAuthentication


class TestGraphModule:
    """Test cases for graph module functions."""

    def setup_method(self):
        """Set up test fixtures before each test method."""
        self.mock_auth = Mock(spec=AzureAuthentication)
        self.mock_auth.get_token.return_value = "mock-access-token"

    def test_set_and_get_auth_instance(self):
        """Test setting and getting the global auth instance."""
        set_auth_instance(self.mock_auth)
        retrieved_auth = get_auth_instance()
        assert retrieved_auth == self.mock_auth

    @patch("src.microsoft_mcp.graph.AzureAuthentication")
    def test_get_auth_instance_creates_default(self, mock_auth_class):
        """Test that get_auth_instance creates a default instance when none exists."""
        # Reset global auth instance
        import src.microsoft_mcp.graph as graph_module

        graph_module._global_auth = None

        mock_auth_instance = Mock()
        mock_auth_class.return_value = mock_auth_instance

        result = get_auth_instance()

        assert result == mock_auth_instance
        mock_auth_class.assert_called_once()

    @patch("src.microsoft_mcp.graph._client")
    def test_request_get_success(self, mock_client):
        """Test successful GET request."""
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.content = b'{"result": "success"}'
        mock_response.json.return_value = {"result": "success"}
        mock_client.request.return_value = mock_response

        result = request("GET", "/test", auth=self.mock_auth)

        assert result == {"result": "success"}
        mock_client.request.assert_called_once()
        call_args = mock_client.request.call_args
        assert call_args[1]["method"] == "GET"
        assert call_args[1]["url"] == f"{BASE_URL}/test"
        assert "Authorization" in call_args[1]["headers"]
        assert call_args[1]["headers"]["Authorization"] == "Bearer mock-access-token"

    @patch("src.microsoft_mcp.graph._client")
    def test_request_post_with_json(self, mock_client):
        """Test successful POST request with JSON payload."""
        mock_response = Mock()
        mock_response.status_code = 201
        mock_response.content = b'{"created": true}'
        mock_response.json.return_value = {"created": True}
        mock_client.request.return_value = mock_response

        payload = {"name": "test"}
        result = request("POST", "/test", json=payload, auth=self.mock_auth)

        assert result == {"created": True}
        call_args = mock_client.request.call_args[1]
        assert call_args["json"] == payload
        assert call_args["headers"]["Content-Type"] == "application/json"

    @patch("src.microsoft_mcp.graph._client")
    def test_request_search_headers(self, mock_client):
        """Test that search requests include proper headers."""
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.content = b'{"value": []}'
        mock_response.json.return_value = {"value": []}
        mock_client.request.return_value = mock_response

        params = {"$search": "test query"}
        request("GET", "/test", params=params, auth=self.mock_auth)

        headers = mock_client.request.call_args[1]["headers"]
        assert headers["Prefer"] == 'outlook.body-content-type="text"'
        assert headers["ConsistencyLevel"] == "eventual"

    @patch("src.microsoft_mcp.graph._client")
    def test_request_rate_limit_retry(self, mock_client):
        """Test that rate limiting (429) triggers retry."""
        # First call returns 429, second call succeeds
        mock_429_response = Mock()
        mock_429_response.status_code = 429
        mock_429_response.headers = {"Retry-After": "1"}

        mock_success_response = Mock()
        mock_success_response.status_code = 200
        mock_success_response.content = b'{"success": true}'
        mock_success_response.json.return_value = {"success": True}

        mock_client.request.side_effect = [mock_429_response, mock_success_response]

        with patch("time.sleep") as mock_sleep:
            result = request("GET", "/test", auth=self.mock_auth)

        assert result == {"success": True}
        assert mock_client.request.call_count == 2
        mock_sleep.assert_called_once_with(1)  # Retry-After value

    @patch("src.microsoft_mcp.graph._client")
    def test_request_server_error_retry(self, mock_client):
        """Test that server errors (5xx) trigger exponential backoff retry."""
        # First call returns 500, second call succeeds
        mock_error_response = Mock()
        mock_error_response.status_code = 500

        mock_success_response = Mock()
        mock_success_response.status_code = 200
        mock_success_response.content = b'{"success": true}'
        mock_success_response.json.return_value = {"success": True}

        mock_client.request.side_effect = [
            httpx.HTTPStatusError(
                "Server error", request=Mock(), response=mock_error_response
            ),
            mock_success_response,
        ]

        with patch("time.sleep") as mock_sleep:
            result = request("GET", "/test", auth=self.mock_auth)

        assert result == {"success": True}
        assert mock_client.request.call_count == 2
        mock_sleep.assert_called_once_with(1)  # 2^0 * 1 = 1 second

    @patch("src.microsoft_mcp.graph._client")
    def test_request_empty_response(self, mock_client):
        """Test handling of empty response."""
        mock_response = Mock()
        mock_response.status_code = 204  # No Content
        mock_response.content = b""
        mock_client.request.return_value = mock_response

        result = request("DELETE", "/test", auth=self.mock_auth)

        assert result is None

    @patch("src.microsoft_mcp.graph.request")
    def test_request_paginated_single_page(self, mock_request):
        """Test paginated request with single page of results."""
        mock_request.return_value = {
            "value": [{"id": "1"}, {"id": "2"}],
            "@odata.nextLink": None,
        }

        results = list(request_paginated("/test", auth=self.mock_auth))

        assert len(results) == 2
        assert results[0]["id"] == "1"
        assert results[1]["id"] == "2"
        mock_request.assert_called_once_with(
            "GET", "/test", params=None, auth=self.mock_auth
        )

    @patch("src.microsoft_mcp.graph.request")
    def test_request_paginated_multiple_pages(self, mock_request):
        """Test paginated request with multiple pages."""
        # First page
        first_page = {
            "value": [{"id": "1"}, {"id": "2"}],
            "@odata.nextLink": f"{BASE_URL}/test?$skip=2",
        }
        # Second page
        second_page = {"value": [{"id": "3"}], "@odata.nextLink": None}

        mock_request.side_effect = [first_page, second_page]

        results = list(request_paginated("/test", auth=self.mock_auth))

        assert len(results) == 3
        assert [r["id"] for r in results] == ["1", "2", "3"]
        assert mock_request.call_count == 2

    @patch("src.microsoft_mcp.graph.request")
    def test_request_paginated_with_limit(self, mock_request):
        """Test paginated request with limit parameter."""
        mock_request.return_value = {
            "value": [{"id": "1"}, {"id": "2"}, {"id": "3"}],
            "@odata.nextLink": None,
        }

        results = list(request_paginated("/test", limit=2, auth=self.mock_auth))

        assert len(results) == 2
        assert [r["id"] for r in results] == ["1", "2"]

    @patch("src.microsoft_mcp.graph.request")
    def test_search_query_success(self, mock_request):
        """Test successful search query."""
        mock_request.return_value = {
            "value": [
                {
                    "hitsContainers": [
                        {
                            "hits": [
                                {"resource": {"id": "1", "subject": "Test email"}},
                                {"resource": {"id": "2", "subject": "Another test"}},
                            ],
                            "moreResultsAvailable": False,
                        }
                    ]
                }
            ]
        }

        results = list(search_query("test", ["message"], auth=self.mock_auth))

        assert len(results) == 2
        assert results[0]["id"] == "1"
        assert results[1]["id"] == "2"

    @patch("src.microsoft_mcp.graph.request")
    def test_search_query_invalid_entity_types(self, mock_request):
        """Test search query with invalid entity types."""
        results = list(search_query("test", ["invalid_type"], auth=self.mock_auth))

        assert len(results) == 0
        mock_request.assert_not_called()

    @patch("src.microsoft_mcp.graph.request")
    def test_search_query_mixed_entity_types(self, mock_request):
        """Test search query with mix of valid and invalid entity types."""
        mock_request.return_value = {
            "value": [
                {
                    "hitsContainers": [
                        {
                            "hits": [{"resource": {"id": "1"}}],
                            "moreResultsAvailable": False,
                        }
                    ]
                }
            ]
        }

        # Mix valid and invalid entity types
        results = list(
            search_query(
                "test", ["message", "invalid_type", "event"], auth=self.mock_auth
            )
        )

        assert len(results) == 1
        # Should only use valid entity types
        call_args = mock_request.call_args[1]["json"]
        assert set(call_args["requests"][0]["entityTypes"]) == {"message", "event"}

    @patch("src.microsoft_mcp.graph.request")
    def test_search_query_with_limit(self, mock_request):
        """Test search query with result limit."""
        mock_request.return_value = {
            "value": [
                {
                    "hitsContainers": [
                        {
                            "hits": [
                                {"resource": {"id": "1"}},
                                {"resource": {"id": "2"}},
                                {"resource": {"id": "3"}},
                            ],
                            "moreResultsAvailable": False,
                        }
                    ]
                }
            ]
        }

        results = list(search_query("test", ["message"], limit=2, auth=self.mock_auth))

        assert len(results) == 2

    @patch("src.microsoft_mcp.graph.request")
    def test_search_query_http_error_handling(self, mock_request):
        """Test search query error handling for different HTTP status codes."""
        # Test 400 Bad Request
        error_400 = httpx.HTTPStatusError(
            "Bad request", request=Mock(), response=Mock(status_code=400)
        )
        mock_request.side_effect = error_400

        with pytest.raises(ValueError, match="Bad request - invalid search query"):
            list(search_query("test", ["message"], auth=self.mock_auth))

        # Test 403 Forbidden
        error_403 = httpx.HTTPStatusError(
            "Forbidden", request=Mock(), response=Mock(status_code=403)
        )
        mock_request.side_effect = error_403

        with pytest.raises(
            PermissionError, match="Forbidden - insufficient permissions"
        ):
            list(search_query("test", ["message"], auth=self.mock_auth))

    @patch("src.microsoft_mcp.graph.request")
    def test_search_query_network_error(self, mock_request):
        """Test search query network error handling."""
        mock_request.side_effect = httpx.RequestError("Network error")

        with pytest.raises(ConnectionError, match="Network error during search query"):
            list(search_query("test", ["message"], auth=self.mock_auth))
