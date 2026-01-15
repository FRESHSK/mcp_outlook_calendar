# Outlook MCP Server

A Model Context Protocol (MCP) server for Microsoft Outlook, built with Node.js.
This server allows AI assistants (like Claude Desktop) to interact with your Outlook inbox and Calendar using the Microsoft Graph API.

## Features

- **Device Code Authentication**: Secure login via Azure AD with persistent token caching.
- **Mail Tools**:
  - `list_messages`: View top 10 emails from your Inbox (Subject, Sender, Preview).
  - `get_message`: Read the full HTML content of a specific email.
  - `send_message`: Send new emails.
- **Calendar Tools**:
  - `list_events`: List events for the next 7 days.
  - `create_event`: Create new events with attendees. **Automatically adds a Microsoft Teams meeting link**.
- **Robustness**: Auto-refresh tokens, handles Graph API errors gracefully.

## Prerequisites

- Node.js (v16+)
- An Azure Cloud account (to create an App Registration)

## Setup

### 1. Azure App Registration

1.  Go to [Azure Portal](https://portal.azure.com/) > **App permissions**.
2.  Create a new registration (e.g., "Outlook MCP").
3.  In **Authentication**, enable "Allow public client flows" (Device Code Flow).
4.  In **API Permissions**, add the following **Delegated** permissions for **Microsoft Graph**:
    - `User.Read`
    - `Mail.Read`
    - `Mail.Send`
    - `Calendars.Read`
    - `Calendars.ReadWrite`
    - (Optional) `Grant admin consent` if using an organization account.

### 2. Installation

```bash
git clone <your-repo-url>
cd mcp_outlook_01
npm install
```

### 3. Configuration

Create a `.env` file in the root directory:

```env
MICROSOFT_CLIENT_ID=your_client_id_from_azure
MICROSOFT_TENANT_ID=common
```
*(Note: Use `common` for personal accounts or your Tenant ID for organization accounts).*

### 4. First Run (Authentication)

Run the server manually once to effectuate the login:

```bash
node server.js
```

Follow the instructions to open `microsoft.com/devicelogin` and enter the code provided.
Once successfull, a `token_cache.json` file will be created. You can verify it works when you see "Outlook MCP Server running on stdio".

## Usage with Claude Desktop

Add the following to your Claude Desktop configuration file (`%APPDATA%\Claude\claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:\\path\\to\\mcp_outlook_01\\server.js"]
    }
  }
}
```

Restart Claude Desktop. You can now ask Claude:
- "Check my latest emails"
- "Read the email from GitHub"
- "Send an email to my colleague"
- "What do I have planned this week?"
- "Schedule a meeting with [email] tomorrow at 10 AM" (will create a Teams meeting)

## Troubleshooting

- **`fetch is not a function`**: Ensure you are using `node-fetch@2` (`npm install node-fetch@2`). This project uses CommonJS.
- **Authentication Errors / New Permissions**: If you change permissions in Azure (e.g. adding Calendar), **delete `token_cache.json`** and run `node server.js` manually to re-authenticate and accept the new scopes.
- **Error sending email**: If you see JSON errors during sending, it might be the 202 Accepted response. Ensure your `server.js` handles status 202 (already patched in this repo).
