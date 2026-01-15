const { McpServer } = require("@modelcontextprotocol/sdk/server/mcp.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const { z } = require("zod");
const msal = require("@azure/msal-node");
const fetch = require("node-fetch");
const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");

// Load .env from absolute path to support execution from any CWD
const envPath = path.resolve(__dirname, ".env");
// Silence console.log during dotenv loading to prevent pollution of stdout (MCP protocol)
const originalLog = console.log;
console.log = () => { };
dotenv.config({ path: envPath });
console.log = originalLog;

const CLIENT_ID = process.env.MICROSOFT_CLIENT_ID;
const TENANT_ID = process.env.MICROSOFT_TENANT_ID || "common";
const TOKEN_CACHE_PATH = path.resolve(__dirname, "token_cache.json");

if (!CLIENT_ID) {
  console.error("Error: MICROSOFT_CLIENT_ID is not defined in .env");
  process.exit(1);
}

// MSAL Configuration
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
  cache: {
    cachePlugin: {
      beforeCacheAccess: async (cacheContext) => {
        if (fs.existsSync(TOKEN_CACHE_PATH)) {
          const data = fs.readFileSync(TOKEN_CACHE_PATH, "utf-8");
          cacheContext.tokenCache.deserialize(data);
        }
      },
      afterCacheAccess: async (cacheContext) => {
        if (cacheContext.cacheHasChanged) {
          const data = cacheContext.tokenCache.serialize();
          fs.writeFileSync(TOKEN_CACHE_PATH, data, "utf-8");
        }
      },
    },
  },
};

const pca = new msal.PublicClientApplication(msalConfig);
const SCOPES = ["User.Read", "Mail.Read", "Mail.Send", "Calendars.Read", "Calendars.ReadWrite"];

// Authentication Helper
async function getToken() {
  const accounts = await pca.getTokenCache().getAllAccounts();

  if (accounts.length > 0) {
    try {
      const silentRequest = {
        account: accounts[0],
        scopes: SCOPES,
      };
      const response = await pca.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (error) {
      console.error("Silent token acquisition failed, falling back to device code.");
    }
  }

  const deviceCodeRequest = {
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      console.error(response.message); // Log to stderr so it doesn't interfere with MCP protocol if using stdio
    },
  };

  const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
  return response.accessToken;
}

// Graph API Helper
async function callGraph(endpoint, options = {}) {
  const token = await getToken();
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    ...options.headers,
  };

  const response = await fetch(endpoint, {
    ...options,
    headers,
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(`Graph API Eror: ${response.status} ${response.statusText} - ${errorBody}`);
  }

  // Handle 204 No Content and 202 Accepted
  if (response.status === 204 || response.status === 202) {
    return null;
  }

  return response.json();
}

// MCP Server Setup
const server = new McpServer({
  name: "outlook-mcp-server",
  version: "1.0.0",
});

server.tool(
  "list_messages",
  "Get the top 10 messages from the Inbox.",
  {},
  async () => {
    try {
      const result = await callGraph("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=10&$select=id,subject,from,receivedDateTime,bodyPreview,isRead");
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result.value, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing messages: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "get_message",
  "Get the full content of a specific message by ID.",
  {
    message_id: z.string().describe("The ID of the message to retrieve"),
  },
  async ({ message_id }) => {
    try {
      const result = await callGraph(`https://graph.microsoft.com/v1.0/me/messages/${message_id}`);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting message: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "list_events",
  "Get calendar events for the next 7 days.",
  {},
  async () => {
    try {
      const start = new Date();
      const end = new Date();
      end.setDate(start.getDate() + 7);

      const result = await callGraph(
        `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}&$select=subject,start,end,location`
      );

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result.value, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing events: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "create_event",
  "Create a new calendar event.",
  {
    subject: z.string().describe("The subject of the event"),
    startDateTime: z.string().describe("Start time in ISO 8601 format (e.g., 2024-01-01T10:00:00)"),
    endDateTime: z.string().describe("End time in ISO 8601 format"),
    location: z.string().optional().describe("Location of the event"),
    content: z.string().optional().describe("Description/body of the event"),
    attendees: z.array(z.string()).optional().describe("List of email addresses to invite"),
  },
  async ({ subject, startDateTime, endDateTime, location, content, attendees }) => {
    console.error(`[create_event] Called with:`, { subject, startDateTime, endDateTime, attendees });
    try {
      const event = {
        subject: subject,
        start: {
          dateTime: startDateTime,
          timeZone: "UTC",
        },
        end: {
          dateTime: endDateTime,
          timeZone: "UTC",
        },
        location: {
          displayName: location || "",
        },
        body: {
          contentType: "Text",
          content: content || "",
        },
        attendees: attendees ? attendees.map(email => ({
          emailAddress: {
            address: email,
            name: email
          },
          type: "required"
        })) : [],
        isOnlineMeeting: true,
        onlineMeetingProvider: "teamsForBusiness",
      };

      console.error(`[create_event] Payload:`, JSON.stringify(event, null, 2));

      const result = await callGraph("https://graph.microsoft.com/v1.0/me/events", {
        method: "POST",
        body: JSON.stringify(event),
      });

      console.error(`[create_event] Success:`, result.id);

      return {
        content: [
          {
            type: "text",
            text: `Event created successfully: ${JSON.stringify(result, null, 2)}`,
          },
        ],
      };
    } catch (error) {
      console.error(`[create_event] Error:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating event: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "send_message",
  "Send an email.",
  {
    subject: z.string().describe("The subject of the email"),
    toRecipients: z.array(z.string()).describe("List of email addresses to send to"),
    content: z.string().describe("The body content of the email"),
    contentType: z.enum(["Text", "HTML"]).optional().default("Text").describe("Content type of the body"),
  },
  async ({ subject, toRecipients, content, contentType }) => {
    try {
      const message = {
        message: {
          subject: subject,
          body: {
            contentType: contentType,
            content: content,
          },
          toRecipients: toRecipients.map(email => ({
            emailAddress: {
              address: email,
            },
          })),
        },
      };

      await callGraph("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        body: JSON.stringify(message),
      });

      return {
        content: [
          {
            type: "text",
            text: "Message sent successfully.",
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error sending message: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

async function main() {
  // Attempt to acquire token on startup to facilitate manual setup/logging
  // If running in Claude with no token, this might time out, but that's why we have the manual setup step.
  try {
    await getToken();
  } catch (err) {
    console.error("Failed to acquire initial token:", err);
    // We don't exit, we let the server start, maybe tools will try again or fail gracefully
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Outlook MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
