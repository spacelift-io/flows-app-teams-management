import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";
import { messageSchema } from "../../schemas/messageSchema.ts";

export const getMessage: AppBlock = {
  name: "Get Message",
  description:
    "Retrieves a specific message from a Teams channel by its ID.\n\n**Required Permission:** ChannelMessage.Read.All",
  category: "Messaging",
  inputs: {
    default: {
      name: "Get Message",
      description: "Retrieve a specific channel message.",
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
          description: "The ID of the message to retrieve.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId, messageId } = input.event.inputConfig;

        const message = await callGraphApi(
          `/teams/${teamId}/channels/${channelId}/messages/${messageId}`,
          accessToken,
        );

        await events.emit(message);
      },
    },
  },
  outputs: {
    default: {
      name: "Message",
      description: "The retrieved message",
      type: messageSchema,
    },
  },
};
