#!/usr/bin/env python3
"""
Authenticate Microsoft account for use with Microsoft MCP.
Run this script to sign in to your Microsoft account using delegated access.
"""

import os
import sys
import asyncio
from pathlib import Path

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / "src"))

from dotenv import load_dotenv
from microsoft_mcp import auth

# Load environment variables before anything else
load_dotenv()


async def main():
    if not os.getenv("MICROSOFT_MCP_CLIENT_ID"):
        print("Error: MICROSOFT_MCP_CLIENT_ID environment variable is required")
        print("\nPlease set it in your .env file or environment:")
        print("export MICROSOFT_MCP_CLIENT_ID='your-app-id'")
        print("\nNote: This should be the Application (client) ID from your")
        print("Azure AD app registration configured for delegated access.")
        print("\nOptional environment variables:")
        print("- MICROSOFT_MCP_TENANT_ID: Tenant ID (defaults to 'common')")
        print(
            "- MICROSOFT_MCP_REDIRECT_URI: Custom redirect URI for non-localhost deployments"
        )
        sys.exit(1)

    print("Microsoft MCP Delegated Access Authentication")
    print("============================================")
    print("This tool will authenticate using delegated access, allowing")
    print("the app to access Microsoft Graph on behalf of the signed-in user.")
    print("Authentication will open a browser window for sign-in.")

    # Show configuration info
    redirect_uri = os.getenv("MICROSOFT_MCP_REDIRECT_URI")
    if redirect_uri:
        print(f"Using custom redirect URI: {redirect_uri}")
    else:
        print("Using default localhost redirect URI")
    print()

    # Check if already authenticated
    try:
        print("Checking current authentication status...")
        user_info = await auth.get_user_info()
        print(f"✓ Already authenticated as: {user_info['displayName']}")
        print(f"  Email: {user_info.get('mail') or user_info.get('userPrincipalName')}")
        print(f"  User ID: {user_info['id']}")
        
        choice = input("\nDo you want to re-authenticate? (y/n): ").lower()
        if choice != "y":
            print("Using existing authentication.")
            return
        else:
            # Clear existing cache to force re-authentication
            auth.clear_token_cache()
            print("Token cache cleared. Proceeding with authentication...")
    except Exception:
        print("No valid authentication found. Proceeding with authentication...")

    print()

    try:
        print("Starting authentication process...")
        print("This will open a browser window for Microsoft sign-in.")
        print("\nRequested permissions:")
        for scope in auth.SCOPES:
            print(f"   - {scope}")
        print("\nStarting authentication...")

        # Trigger authentication by trying to get user info
        user_info = await auth.get_user_info()

        print("\n✓ Authentication successful!")
        print(f"Signed in as: {user_info['displayName']}")
        print(f"Email: {user_info.get('mail') or user_info.get('userPrincipalName')}")
        print(f"User ID: {user_info['id']}")
        print("✓ Delegated access verified")

    except Exception as e:
        print(f"\n✗ Authentication failed: {e}")
        sys.exit(1)

    print("\nDelegated Access Permissions:")
    print("The authenticated account has consented to the following permissions:")
    print("• User.Read - Read user profile")
    print("• User.ReadBasic.All - Read basic info of all users")
    print("• Mail.Read - Read emails")
    print("• Team.ReadBasic.All - Read basic team information")
    print("• TeamMember.ReadWrite.All - Read and write team membership")
    print("• Calendars.Read - Access calendars")
    print("• Files.Read - Access OneDrive files")

    print("\n✓ Delegated Access Authentication complete!")
    print("You can now use the Microsoft MCP tools without specifying account_id.")


if __name__ == "__main__":
    asyncio.run(main())
