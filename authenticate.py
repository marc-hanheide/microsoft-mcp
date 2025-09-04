#!/usr/bin/env python3
"""
Authenticate Microsoft account for use with Microsoft MCP.
Run this script to sign in to your Microsoft account using delegated access.
"""

import os
import sys
from pathlib import Path

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / "src"))

from dotenv import load_dotenv
from microsoft_mcp.auth import AzureAuthentication
from microsoft_mcp import graph

# Load environment variables before anything else
load_dotenv()


def main():
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

    # Get auth instance
    auth = AzureAuthentication(
        auth_record_file=os.getenv("AZURE_CRED_CACHE_FILE"),
        token_cache_file=os.getenv("AZURE_TOKEN_CACHE_FILE")
    )

    # Set the auth instance for the graph module
    graph.set_auth_instance(auth)

    # Check if already authenticated
    try:
        print("Checking current authentication status...")

        # Check if we have an AuthenticationRecord and can get a token
        if auth.exists_valid_token():
            # Try to get user info to verify authentication works
            user_info = graph.request(
                "GET",
                "/me",
                params={"$select": "id,displayName,mail,userPrincipalName"},
            )

            print(f"✓ Already authenticated as: {user_info['displayName']}")
            print(
                f"  Email: {user_info.get('mail') or user_info.get('userPrincipalName')}"
            )
            print(f"  User ID: {user_info['id']}")

            # Display current token information
            try:
                import datetime

                token, expires_on = auth.get_token_with_details()
                expires_dt = datetime.datetime.fromtimestamp(expires_on)

                print(f"\n📋 Current Token Information:")
                print(f"   Token (first 20 chars): {token[:20]}...")
                print(f"   Expires on: {expires_dt.strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"   Expires in: {expires_dt - datetime.datetime.now()}")
            except Exception as e:
                print(f"   ⚠ Could not retrieve token details: {e}")

            choice = input("\nDo you want to re-authenticate? (y/n): ").lower()
            if choice != "y":
                print("Using existing authentication.")
                return
            else:
                # Clear existing cache to force re-authentication
                auth.clear_cache()
                print("Authentication cache cleared. Proceeding with authentication...")
        else:
            print("No valid authentication found. Proceeding with authentication...")

    except Exception as e:
        print(f"Authentication check failed: {e}")
        print("Proceeding with authentication...")

    print()

    try:
        print("Starting authentication process...")
        print("This will open a browser window for Microsoft sign-in.")
        print("\nRequested permissions:")
        from microsoft_mcp.auth import SCOPES

        for scope in SCOPES:
            print(f"   - {scope}")
        print("\nStarting authentication...")

        # Perform interactive authentication
        auth_record = auth.authenticate()
        print(f"\n✓ Authentication successful!")
        print(f"AuthenticationRecord saved to: {auth.auth_record_file}")

        # Verify authentication by getting user info
        user_info = graph.request(
            "GET", "/me", params={"$select": "id,displayName,mail,userPrincipalName"}
        )

        print(f"Signed in as: {user_info['displayName']}")
        print(f"Email: {user_info.get('mail') or user_info.get('userPrincipalName')}")
        print(f"User ID: {user_info['id']}")
        print("✓ Delegated access verified")

        # Get and display token information
        try:
            import datetime

            token, expires_on = auth.get_token_with_details()
            expires_dt = datetime.datetime.fromtimestamp(expires_on)

            print(f"\n📋 Token Information:")
            print(f"   Token (first 20 chars): {token[:20]}...")
            print(f"   Expires on: {expires_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"   Expires in: {expires_dt - datetime.datetime.now()}")
        except Exception as e:
            print(f"⚠ Could not retrieve token details: {e}")

    except Exception as e:
        print(f"\n✗ Authentication failed: {e}")
        sys.exit(1)

    print("\nDelegated Access Permissions:")
    print("The authenticated account has consented to the following permissions:")
    print("• User.Read - Read user profile")
    print("• User.ReadBasic.All - Read basic info of all users")
    print("• Chat.Read - Read chat messages")
    print("• Mail.Read - Read emails")
    print("• Team.ReadBasic.All - Read basic team information")
    print("• TeamMember.ReadWrite.All - Read and write team membership")
    print("• Calendars.Read - Access calendars")
    print("• Files.Read - Access OneDrive files")

    print("\n✓ Delegated Access Authentication complete!")
    print("You can now use the Microsoft MCP tools.")
    print(
        "Future runs will authenticate silently using the saved AuthenticationRecord."
    )


if __name__ == "__main__":
    main()
