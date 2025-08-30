import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";
import { messageSchema } from "../../schemas/messageSchema.ts";

export const listReplies: AppBlock = {
  name: "List Replies",
  description:
    "Lists all replies to a specific message in a Teams channel. Replies are threaded responses to a message.\n\n**Required Permission:** ChannelMessage.Read.All",
  category: "Messaging",
  inputs: {
    default: {
      name: "List Replies",
      description: "Retrieve all replies to a message.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team containing the channel.",
          type: "string",
          required: true,
        },
        channelId: {
          name: "Channel ID",
          description: "The ID of the channel containing the message.",
          type: "string",
          required: true,
        },
        messageId: {
          name: "Message ID",
          description: "The ID of the message to get replies for.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId, messageId } = input.event.inputConfig;

        const response = await callGraphApi(
          `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`,
          accessToken,
        );

        await events.emit({ replies: response.value || [] });
      },
    },
  },
  outputs: {
    default: {
      name: "Replies",
      description: "List of replies to the message",
      type: {
        type: "object",
        properties: {
          replies: {
            type: "array",
            description: "Array of reply messages",
            items: messageSchema,
          },
        },
        required: ["replies"],
      },
    },
  },
};
