#!/usr/bin/env python3
"""
Authenticate Microsoft accounts for use with Microsoft MCP.
Run this script to sign in to one or more Microsoft accounts.
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
    print("the app to access Microsoft Graph on behalf of signed-in users.")
    print("Authentication will open a browser window for sign-in.")

    # Show configuration info
    redirect_uri = os.getenv("MICROSOFT_MCP_REDIRECT_URI")
    if redirect_uri:
        print(f"Using custom redirect URI: {redirect_uri}")
    else:
        print("Using default localhost redirect URI")
    print()

    # List current accounts
    accounts = await auth.list_accounts_async()
    if accounts:
        print("Currently authenticated accounts:")
        for i, account in enumerate(accounts, 1):
            print(f"{i}. {account.username} (ID: {account.account_id})")
        print()
    else:
        print("No accounts currently authenticated.\n")

    # Authenticate new account
    while True:
        choice = input("Do you want to authenticate a new account? (y/n): ").lower()
        if choice == "n":
            break
        elif choice == "y":
            try:
                # Use the new authentication function
                new_account = await auth.authenticate_new_account()

                if new_account:
                    print("\n✓ Authentication successful!")
                    print(f"Signed in as: {new_account.username}")
                    print(f"Account ID: {new_account.account_id}")
                    print("✓ Delegated access verified during authentication")
                else:
                    print(
                        "\n✗ Authentication failed: Could not retrieve account information"
                    )
            except Exception as e:
                print(f"\n✗ Authentication failed: {e}")
                continue

            print()
        else:
            print("Please enter 'y' or 'n'")

    # Final account summary
    accounts = await auth.list_accounts_async()
    if accounts:
        print("\nAuthenticated accounts summary:")
        print("==============================")
        for account in accounts:
            print(f"• {account.username}")
            print(f"  Account ID: {account.account_id}")

        print(
            "\nYou can use these account IDs with any MCP tool by passing account_id parameter."
        )
        print("Example: send_email(..., account_id='<account-id>')")
        print("\nDelegated Access Permissions:")
        print("The authenticated accounts have consented to the following permissions:")
        print("• User.Read - Read user profile")
        print("• User.ReadBasic.All - Read basic info of all users")
        print("• Mail.Read/Send - Read and send emails")
        print("• Team.ReadBasic.All - Read basic team information")
        print("• TeamMember.ReadWrite.All - Read and write team membership")
        print("• Calendars.Read - Access calendars")
        print("• Files.Read - Access OneDrive files")
    else:
        print("\nNo accounts authenticated.")

    print("\nDelegated Access Authentication complete!")


if __name__ == "__main__":
    asyncio.run(main())
