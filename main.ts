import { defineApp, http, kv } from "@slflows/sdk/v1";
import { blocks } from "./blocks/index.ts";
import {
  refreshAccessToken,
  TOKEN_KV_KEY,
  TOKEN_EXPIRY_KV_KEY,
  ensureCentralSubscription,
  deleteSubscription,
} from "./utils/teamsClient.ts";
import { handleLifecycleNotification } from "./utils/handleLifecycleNotification.ts";
import { handleWebhookNotification } from "./utils/handleWebhookNotification.ts";

export const app = defineApp({
  name: "Teams Management",
  installationInstructions: `
# Microsoft Teams Management App Setup Guide

This app provides monitoring and administrative capabilities for Microsoft Teams, including message subscriptions, channel management, and member administration. It uses application-only authentication for server-to-server access.

**Note:** This app can read and manage Teams resources but cannot send messages. For interactive messaging, use a separate Teams Bot application.

To connect your Microsoft Teams workspace, you'll need to register an application in Microsoft Entra ID (formerly Azure AD) and configure it with the required permissions.

## Step 1: Register Application in Microsoft Entra ID

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** (or **Azure Active Directory**)
3. Select **Add** → **App registration**
4. Fill in the registration form:
   - **Name**: "Spacelift Flows - Teams Integration" (or your preferred name)
   - **Supported account types**: Select "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank for now
   - Click **Register**

## Step 2: Create Client Secret

1. In your new app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description (e.g., "Flows Integration Secret")
4. Choose an expiration period (recommended: 24 months)
5. Click **Add**
6. **IMPORTANT**: Copy the **Value** immediately - you won't be able to see it again
7. Save this as your **Client Secret** below

## Step 3: Copy Application IDs

1. Go to the **Overview** page of your app registration
2. Copy the **Application (client) ID** - this is your **Client ID**
3. Copy the **Directory (tenant) ID** - this is your **Tenant ID**

## Step 4: Configure API Permissions

1. Go to **API permissions** in your app registration
2. Click **Add a permission** → **Microsoft Graph** → **Application permissions**
3. Add the required permissions based on which blocks you plan to use:

### Permission Reference by Feature

**Message Operations** (getMessage, listReplies, messagesSubscription):
- **ChannelMessage.Read.All** - Read all channel messages and their replies

**Channel Information** (listChannels, getChannelInfo):
- **Channel.ReadBasic.All** - Read basic channel information (names, IDs)
- **Channel.Read.All** - Read detailed channel information and settings

**Channel Management** (createChannel, archiveChannel, unarchiveChannel):
- **Channel.ReadWrite.All** - Create, update, and archive channels

**Channel Membership** (listChannelMembers, inviteUsersToChannel, removeUserFromChannel):
- **ChannelMember.Read.All** - List channel members
- **ChannelMember.ReadWrite.All** - Add and remove channel members

**Team Operations** (listTeams, addTeamMember):
- **Group.Read.All** - List teams and read team information
- **GroupMember.ReadWrite.All** - Add and remove team members

**User Operations** (getUser, listUsers):
- **User.Read.All** - Read user profile information

### Recommended Minimal Setup

For basic monitoring and read operations:
- Channel.ReadBasic.All
- ChannelMember.Read.All
- ChannelMessage.Read.All
- Group.Read.All
- User.Read.All

For full administrative capabilities, add:
- Channel.ReadWrite.All
- ChannelMember.ReadWrite.All
- GroupMember.ReadWrite.All

4. Click **Grant admin consent** for your organization
   - This requires Global Administrator or Privileged Role Administrator permissions
   - Admin consent is required for all application permissions

**Note:** This app uses application-only authentication (client credentials flow) which allows reading and managing Teams resources, but **cannot send messages**. For interactive messaging capabilities, a separate Teams Bot app is required.

## Step 5: Configure Webhooks (After Initial Setup)

After you've configured the app with your credentials below, you'll receive a webhook URL. You'll need to:

1. Use this webhook URL when creating Graph API change notification subscriptions
2. The app will automatically handle webhook validation and renewal

## Step 6: Complete Configuration

Fill in the fields below with the values you copied:
- **Client ID**: Application (client) ID from Step 3
- **Client Secret**: Secret value from Step 2
- **Tenant ID**: Directory (tenant) ID from Step 3
- **Enable Message Subscriptions**: Check this if you want to use subscription blocks (requires additional permissions from Step 4)

Click **Confirm** to complete the setup.

## Troubleshooting

- **Authentication fails**: Verify your Client Secret is correct and hasn't expired
- **Permission errors**: Ensure admin consent was granted for all required permissions
- **Webhook subscription fails**: Check that your webhook URL is accessible from Microsoft's servers

## Additional Resources

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/overview)
- [Teams Change Notifications](https://learn.microsoft.com/en-us/graph/teams-changenotifications-chatmessage)
- [Adaptive Cards Designer](https://adaptivecards.io/designer/)
`,

  config: {
    tenantId: {
      name: "Tenant ID (Directory ID)",
      description:
        "The Directory (tenant) ID from your Azure AD app registration.",
      type: "string",
      sensitive: false,
      required: true,
    },
    clientId: {
      name: "Client ID (Application ID)",
      description:
        "The Application (client) ID from your Azure AD app registration.",
      type: "string",
      sensitive: false,
      required: true,
    },
    clientSecret: {
      name: "Client Secret",
      description:
        "The client secret value created in your Azure AD app registration.",
      type: "string",
      sensitive: true,
      required: true,
    },
    enableSubscriptions: {
      name: "Enable Message Subscriptions",
      description:
        "Enable webhook subscriptions for Teams messages and events. Requires Chat.Read.All or Chat.ReadWrite.All permission. Disable if only using channel management features.",
      type: "boolean",
      sensitive: false,
      required: false,
      default: false,
    },
  },

  signals: {
    subscriptionId: {
      name: "Subscription ID",
      description:
        "The Microsoft Graph API subscription ID for channel messages",
    },
    subscriptionExpiry: {
      name: "Subscription Expiry",
      description: "When the subscription expires (timestamp)",
    },
  },

  schedules: {
    tokenRefresh: {
      description: "Refresh Microsoft Graph access token every 50 minutes",
      definition: {
        type: "frequency",
        frequency: {
          interval: 50,
          unit: "minutes",
        },
      },
      async onTrigger(input) {
        // Refresh access token every 50 minutes (tokens last ~1 hour)
        try {
          const { accessToken, tokenExpiry } = await refreshAccessToken(
            input.app.config,
          );

          // Store in app KV
          await kv.app.set({ key: TOKEN_KV_KEY, value: accessToken });
          await kv.app.set({
            key: TOKEN_EXPIRY_KV_KEY,
            value: tokenExpiry.toString(),
          });

          console.log("Token refreshed successfully");
        } catch (error: any) {
          console.error("Failed to refresh token on schedule:", error.message);
          // Don't throw - let the next schedule run retry
        }
      },
    },
  },

  async onSync(input) {
    const { clientId, clientSecret, tenantId } = input.app.config;

    try {
      // Get access token from Microsoft Graph
      const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      const tokenParams = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      });

      const tokenResponse = await fetch(tokenUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: tokenParams.toString(),
      });

      if (!tokenResponse.ok) {
        const errorData = await tokenResponse.json();
        console.error("Token acquisition failed:", errorData);
        return {
          newStatus: "failed",
          customStatusDescription: `Authentication failed, see logs`,
        };
      }

      const tokenData = await tokenResponse.json();

      // Calculate token expiry time
      const expiresIn = tokenData.expires_in || 3599;
      const tokenExpiry = Date.now() + expiresIn * 1000;

      // Verify token by calling Graph API
      const graphResponse = await fetch(
        "https://graph.microsoft.com/v1.0/organization",
        {
          headers: {
            Authorization: `Bearer ${tokenData.access_token}`,
          },
        },
      );

      if (!graphResponse.ok) {
        const errorData = await graphResponse.json();
        console.error("Graph API verification failed:", errorData);
        return {
          newStatus: "failed",
          customStatusDescription: "API access failed, see logs",
        };
      }

      // Store token in app KV (not in signals)
      await kv.app.set({ key: TOKEN_KV_KEY, value: tokenData.access_token });
      await kv.app.set({
        key: TOKEN_EXPIRY_KV_KEY,
        value: tokenExpiry.toString(),
      });

      // Create central subscription for all Teams communications (if enabled)
      if (input.app.config.enableSubscriptions) {
        try {
          const webhookUrl = `${input.app.http.url}/webhook`;
          const lifecycleUrl = `${input.app.http.url}/lifecycle`;
          const subscriptionInfo = await ensureCentralSubscription(
            webhookUrl,
            lifecycleUrl,
            input.app.config,
            input.app.signals,
          );

          if (subscriptionInfo.action === "created") {
            console.log("Subscription created for webhook:", webhookUrl);
            console.log("Lifecycle notifications at:", lifecycleUrl);
          } else if (subscriptionInfo.action === "renewed") {
            console.log("Subscription renewed successfully");
          }
          // action === "valid": no log needed, already logged in teamsClient

          return {
            newStatus: "ready",
            signalUpdates: {
              subscriptionId: subscriptionInfo.subscriptionId,
              subscriptionExpiry: subscriptionInfo.subscriptionExpiry,
            },
          };
        } catch (subscriptionError: any) {
          console.error(
            "Failed to create subscription:",
            subscriptionError.message,
          );
          return {
            newStatus: "failed",
            customStatusDescription: "Subscription failed, see logs",
          };
        }
      } else {
        console.log(
          "Message subscriptions disabled - subscription blocks will not receive events",
        );

        // Clean up existing subscription if it exists
        const existingSubscriptionId = input.app.signals.subscriptionId;
        if (existingSubscriptionId) {
          try {
            const { accessToken } = await refreshAccessToken(input.app.config);
            await deleteSubscription(existingSubscriptionId, accessToken);
            console.log(
              "Deleted existing subscription:",
              existingSubscriptionId,
            );

            // Clear signals only after successful deletion
            return {
              newStatus: "ready",
              signalUpdates: {
                subscriptionId: null,
                subscriptionExpiry: null,
              },
            };
          } catch (deleteError: any) {
            console.error(
              "Failed to delete existing subscription:",
              deleteError.message,
            );
            return {
              newStatus: "failed",
              customStatusDescription:
                "Failed to delete subscription, see logs",
            };
          }
        }
      }

      return { newStatus: "ready" };
    } catch (error: any) {
      console.error("Error during Teams authentication:", error);
      return {
        newStatus: "failed",
        customStatusDescription: "Setup error, see logs",
      };
    }
  },

  http: {
    async onRequest(input) {
      const requestPath = input.request.path;

      // Handle lifecycle notifications (reauthorization, subscription removed)
      if (requestPath === "/lifecycle" || requestPath.endsWith("/lifecycle")) {
        await handleLifecycleNotification(input);
        return;
      }

      // Handle data notifications
      if (requestPath === "/webhook" || requestPath.endsWith("/webhook")) {
        await handleWebhookNotification(input);
        return;
      }

      // Unknown endpoint
      console.warn("Received request on unhandled HTTP path:", requestPath);
      await http.respond(input.request.requestId, {
        statusCode: 404,
        body: { error: "Endpoint not found" },
      });
    },
  },

  async onDrain(input) {
    // Clean up subscription when app is being drained
    const existingSubscriptionId = input.app.signals.subscriptionId;
    if (existingSubscriptionId) {
      try {
        const { accessToken } = await refreshAccessToken(input.app.config);
        await deleteSubscription(existingSubscriptionId, accessToken);
        console.log(
          "Deleted subscription during drain:",
          existingSubscriptionId,
        );
      } catch (error: any) {
        console.error(
          "Failed to delete subscription during drain:",
          error.message,
        );
        // Don't throw - allow drain to complete
      }
    }

    return {};
  },

  blocks,
});
