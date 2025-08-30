/**
 * Base message properties shared between all message schemas
 */
const baseMessageProperties = {
  id: {
    type: "string",
    description: "The unique ID of the message",
  },
  replyToId: {
    type: "string",
    description: "ID of the parent message (for replies)",
  },
  etag: {
    type: "string",
    description: "Version identifier for the message",
  },
  messageType: {
    type: "string",
    description:
      "Type of message (message, systemEventMessage, unknownFutureValue)",
  },
  createdDateTime: {
    type: "string",
    description: "When the message was created (ISO 8601 format)",
  },
  lastModifiedDateTime: {
    type: "string",
    description: "When the message was last modified (ISO 8601 format)",
  },
  lastEditedDateTime: {
    type: "string",
    description: "When the message was last edited by the user",
  },
  deletedDateTime: {
    type: "string",
    description: "When the message was deleted",
  },
  subject: {
    type: "string",
    description: "The subject of the message",
  },
  summary: {
    type: "string",
    description: "Summary text of the message for notifications",
  },
  chatId: {
    type: "string",
    description: "ID of the chat (if in a chat context)",
  },
  importance: {
    type: "string",
    description: "Importance of the message (normal, high, urgent)",
  },
  locale: {
    type: "string",
    description: "Locale of the message",
  },
  webUrl: {
    type: "string",
    description: "Link to the message in Teams",
  },
  channelIdentity: {
    type: "object",
    description: "Identity of the channel where the message was posted",
    properties: {
      teamId: { type: "string", description: "Team ID" },
      channelId: { type: "string", description: "Channel ID" },
    },
  },
  policyViolation: {
    type: "object",
    description: "Policy violation information if applicable",
  },
  eventDetail: {
    type: "object",
    description: "Event details for system messages",
  },
  from: {
    type: "object",
    description: "Information about who sent the message",
    properties: {
      application: {
        type: "object",
        description: "Application that sent the message",
      },
      device: {
        type: "object",
        description: "Device that sent the message",
      },
      user: {
        type: "object",
        description: "User who sent the message",
        properties: {
          id: { type: "string", description: "User ID" },
          displayName: { type: "string", description: "Display name" },
          userIdentityType: {
            type: "string",
            description: "Type of user identity",
          },
          tenantId: { type: "string", description: "Tenant ID" },
        },
      },
      conversation: {
        type: "object",
        description: "Conversation context",
      },
    },
  },
  body: {
    type: "object",
    description: "The content of the message",
    properties: {
      contentType: {
        type: "string",
        description: "Content type (html, text)",
      },
      content: { type: "string", description: "Message content" },
    },
  },
  attachments: {
    type: "array",
    description: "Attachments in the message",
    items: {
      type: "object",
      properties: {
        id: { type: "string", description: "Attachment ID" },
        contentType: {
          type: "string",
          description: "MIME type of the attachment",
        },
        contentUrl: {
          type: "string",
          description: "URL to the attachment content",
        },
        content: {
          type: "string",
          description: "Attachment content (for inline attachments)",
        },
        name: { type: "string", description: "Name of the attachment" },
        thumbnailUrl: {
          type: "string",
          description: "URL to a thumbnail",
        },
        teamsAppId: {
          type: "string",
          description: "ID of the Teams app (for app attachments)",
        },
      },
    },
  },
  mentions: {
    type: "array",
    description: "User/channel mentions in the message",
    items: {
      type: "object",
      properties: {
        id: { type: "number", description: "Mention ID" },
        mentionText: {
          type: "string",
          description: "Text used for the mention",
        },
        mentioned: {
          type: "object",
          description: "Entity that was mentioned",
        },
      },
    },
  },
  reactions: {
    type: "array",
    description: "Reactions to the message",
    items: {
      type: "object",
      properties: {
        reactionType: {
          type: "string",
          description: "Type of reaction (like, heart, etc.)",
        },
        createdDateTime: {
          type: "string",
          description: "When the reaction was created",
        },
        user: {
          type: "object",
          description: "User who reacted",
        },
      },
    },
  },
};

/**
 * Schema for a message retrieved via getMessage block
 * Does not include changeType
 */
export const messageSchema = {
  type: "object",
  description: "Teams channel message",
  properties: baseMessageProperties,
  required: ["id", "createdDateTime", "from", "body"],
};

/**
 * Schema for a message event from messagesSubscription
 * Includes changeType to indicate what triggered the event
 */
export const messageEventSchema = {
  type: "object",
  description: "Teams channel message event",
  properties: {
    ...baseMessageProperties,
    changeType: {
      type: "string",
      description:
        "Type of change that triggered this event (created, updated, deleted)",
    },
  },
  required: [...messageSchema.required, "changeType"],
};
