"""Microsoft MCP - Model Context Protocol server for Microsoft Graph API integration."""

from .auth import (
    get_token,
    get_graph_client,
    clear_token_cache,
    clear_credential_cache,
    exists_valid_token,
    start_token_refresh_service,
    stop_token_refresh_service,
    is_token_refresh_service_running,
)

from .server import main as server_main

__all__ = [
    "get_token",
    "get_graph_client", 
    "clear_token_cache",
    "clear_credential_cache",
    "exists_valid_token",
    "start_token_refresh_service",
    "stop_token_refresh_service",
    "is_token_refresh_service_running",
    "server_main",
]

def main() -> None:
    print("Hello from microsoft-mcp!")
