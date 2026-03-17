
import os
import sys
import time
from pathlib import Path
from msal import PublicClientApplication, ConfidentialClientApplication

def get_input(prompt):
    """Get user input."""
    return input(f"> {prompt}: ").strip()

def save_env_file(config, account_name=None):
    """Saves the configuration to a .env file."""
    env_file = Path(".env")
    lines = []
    
    # Read existing lines if file exists, to preserve other settings
    if env_file.exists():
        with open(env_file, "r") as f:
            for line in f:
                # Filter out old settings for the account we are configuring
                is_old_setting = False
                if account_name:
                    if line.startswith(f"ACCOUNT_{account_name.upper()}_"):
                        is_old_setting = True
                elif line.startswith("AZURE_"):
                    is_old_setting = True
                
                if not is_old_setting:
                    lines.append(line.strip())

    # Add new settings
    for key, value in config.items():
        prefix = f"ACCOUNT_{account_name.upper()}_" if account_name else "AZURE_"
        lines.append(f'{prefix}{key}="{value}"')

    with open(env_file, "w") as f:
        f.write("\n".join(lines) + "\n")
    
    print(f"\n✅ Configuration saved to {env_file.resolve()}")

def device_code_flow(client_id, tenant_id):
    """Initiates a device code flow and returns the account."""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = PublicClientApplication(client_id, authority=authority)
    
    flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])

    if "user_code" not in flow:
        raise ValueError("Failed to create device flow. Please check your app registration settings.", flow)

    print("\n" + "="*60)
    print("🚀 ACTION REQUIRED 🚀")
    print(flow["message"])
    print("="*60 + "\n")

    # Poll for token
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        print("✅ Successfully authenticated!")
        return result
    else:
        print("❌ Authentication failed or was cancelled.")
        print(f"Error: {result.get('error')}")
        print(f"Description: {result.get('error_description')}")
        sys.exit(1)

def main():
    print("--- Outlook MCP Server Interactive Setup ---")
    
    account_name_input = get_input("Enter a name for this account (e.g., 'personal', 'work', or leave blank for 'default')")
    account_name = account_name_input if account_name_input else None
    
    print("\nWhat type of permissions did you configure in Azure AD?")
    print("1. Delegated (For personal use, no admin consent required)")
    print("2. Application (For server-side use, requires admin consent)")
    auth_type = get_input("Choose (1 or 2)")

    if auth_type not in ["1", "2"]:
        print("Invalid choice. Exiting.")
        sys.exit(1)

    tenant_id = get_input("Enter your Azure Tenant ID")
    client_id = get_input("Enter your Azure Application (Client) ID")
    
    config = {
        "TENANT_ID": tenant_id,
        "CLIENT_ID": client_id,
    }

    if auth_type == "2": # Application
        client_secret = get_input("Enter your Azure Client Secret")
        config["CLIENT_SECRET"] = client_secret
        print("\nVerifying credentials...")
        # Simple verification
        app = ConfidentialClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}", client_credential=client_secret)
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
             print("✅ Application credentials verified successfully.")
        else:
            print("❌ Credential verification failed. Please check your Tenant ID, Client ID, and Client Secret.")
            print(f"Error: {result.get('error')}")
            print(f"Description: {result.get('error_description')}")
            sys.exit(1)
        
        save_env_file(config, account_name)

    else: # Delegated
        print("\nStarting Device Code authentication flow...")
        time.sleep(1)
        device_code_flow(client_id, tenant_id)
        # For delegated auth via device flow, we only need to store the IDs.
        # The server will handle caching the refresh token on first use.
        save_env_file(config, account_name)

if __name__ == "__main__":
    main()

