"""Microsoft MCP - Model Context Protocol server for Microsoft Graph API integration."""

from .auth import AzureAuthentication
from .server import main as server_main

__all__ = [
    "AzureAuthentication",
    "server_main",
]


def main() -> None:
    print("Hello from microsoft-mcp!")
