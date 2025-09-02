import os
import sys
import asyncio
from .tools import mcp
from . import auth


def main() -> None:
    if not os.getenv("MICROSOFT_MCP_CLIENT_ID"):
        print(
            "Error: MICROSOFT_MCP_CLIENT_ID environment variable is required",
            file=sys.stderr,
        )
        sys.exit(1)

    # Initiate authentication flow at startup
    try:
        print("Initializing Microsoft Graph authentication...", file=sys.stderr)

        # Try to get a token to trigger authentication if needed
        # This will use cached token if available, or prompt for authentication
        token = auth.get_token()

        print("✓ Authentication successful - MCP server starting...", file=sys.stderr)

    except Exception as e:
        print(f"Authentication failed: {e}", file=sys.stderr)
        print(
            "Please run the authenticate.py script first to set up authentication.",
            file=sys.stderr,
        )
        sys.exit(1)

    mcp.run()


if __name__ == "__main__":
    main()
