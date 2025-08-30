# Microsoft Teams - Management

A Flows app for managing Microsoft Teams workspaces. Enables automation workflows for Teams administration, channel management, user operations, and message monitoring.

## What This App Is For

This app is designed for **Teams administration and monitoring scenarios**:

- **Monitor channel activity** - Subscribe to message events across your Teams channels
- **Automate channel management** - Create, archive, and configure channels programmatically
- **Manage team membership** - Add and remove users from teams and channels
- **Audit and compliance** - Track messages, retrieve conversation threads, and monitor team activity
- **User management** - Query user information and team memberships

**Use cases:**

- Automated compliance monitoring and archiving
- Channel provisioning workflows (e.g., create standardized channels for new projects)
- Activity-based channel lifecycle management (archive inactive channels)
- User onboarding/offboarding automation
- Integration with external systems for reporting and analytics

## What This App Is NOT For

This app uses **application-only authentication** which provides administrative access but has limitations:

- ❌ **Interactive messaging** - Cannot send messages, replies, or reactions to users
- ❌ **Bot-driven conversations** - Cannot respond to user mentions or participate in conversations
- ❌ **Real-time chat interactions** - Cannot provide interactive chat experiences

**For interactive messaging scenarios**, use the **Microsoft Teams Bot app** instead, which:

- Can send messages and replies to channels and chats
- Responds to @mentions and user interactions
- Provides conversational bot experiences
- Uses the Bot Framework for real-time interactions

The two apps complement each other: this app handles administrative tasks, while the Bot app handles user-facing interactions.

## Key Features

### Message Monitoring

- **Messages Subscription** - Subscribe to channel message events (created, updated, deleted)
- **Get Message** - Retrieve specific messages by ID
- **List Replies** - Get all replies to a message thread

### Channel Operations

- **List Channels** - Enumerate channels in a team
- **Get Channel Info** - Retrieve channel details and settings
- **Create Channel** - Create new public or private channels
- **Archive/Unarchive Channel** - Manage channel lifecycle

### Membership Management

- **List Channel Members** - View channel membership
- **Invite Users to Channel** - Add users to private channels
- **Remove User from Channel** - Remove users from channels
- **Add Team Member** - Add users to teams

### Team & User Operations

- **List Teams** - Get all teams in your organization
- **Get User** - Retrieve user profile information
- **List Users** - Query users with filtering and search

## Setup

The app provides detailed installation instructions when you add it to your workspace. You'll need:

1. **Microsoft Entra ID Admin Access** - To register the application
2. **Client Credentials** - Application ID, Client Secret, and Tenant ID
3. **API Permissions** - Grant appropriate Microsoft Graph permissions based on features you'll use
4. **Admin Consent** - A Global Administrator must approve the permissions

The app handles authentication, token management, and webhook subscriptions automatically.

## Permission Model

Permissions are granted based on the features you need:

- **Read-only monitoring** - Minimal permissions for viewing messages and channel info
- **Channel management** - Additional permissions to create and configure channels
- **Membership management** - Permissions to add/remove users from teams and channels

See the installation guide for detailed permission requirements per feature.

## Architecture

- **Authentication**: OAuth 2.0 client credentials flow (application-only)
- **API**: Microsoft Graph API v1.0
- **Webhooks**: Automatic subscription management for message events
- **Token Management**: Automatic refresh and secure storage

## When to Use Teams Bot App

Use the separate [**Teams Bot app**](https://github.com/spacelift-io/flows-app-teams-bot) if you need to:

- Send messages to users or channels
- Respond to @mentions in Teams
- Provide interactive conversational experiences
- React to messages or post adaptive cards
- Handle user-initiated bot commands

The Bot app uses the Bot Framework and delegated permissions, enabling user-facing interactions that this management app cannot provide.

## Development

```bash
npm install
npm run typecheck
npm run format
npm run bundle
```

## Resources

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/overview)
- [Teams API Reference](https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview)
- [Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
