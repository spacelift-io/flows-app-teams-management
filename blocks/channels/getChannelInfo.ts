import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const getChannelInfo: AppBlock = {
  name: "Get Channel Info",
  description:
    "Gets information about a Teams channel.\n\n**Required Permission:** Channel.Read.All",
  category: "Channels",
  inputs: {
    default: {
      name: "Get",
      description: "Trigger getting information about the specified channel.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team containing the channel.",
          type: "string",
          required: true,
        },
        channelId: {
          name: "Channel ID",
          description: "The ID of the channel to get information about.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId } = input.event.inputConfig;

        const channelInfo = await callGraphApi(
          `/teams/${teamId}/channels/${channelId}`,
          accessToken,
        );

        await events.emit({
          channel: channelInfo,
        });
      },
    },
  },
  outputs: {
    default: {
      name: "Channel Info Retrieved",
      description:
        "Emitted when channel information has been successfully retrieved.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          channel: {
            type: "object",
            description: "The channel object with detailed information.",
            properties: {
              id: { type: "string", description: "The channel ID" },
              displayName: { type: "string", description: "Channel name" },
              description: {
                type: "string",
                description: "Channel description",
              },
              email: {
                type: "string",
                description: "Email address for the channel",
              },
              webUrl: { type: "string", description: "Web URL to the channel" },
              membershipType: {
                type: "string",
                description: "standard, private, or shared",
              },
              createdDateTime: {
                type: "string",
                description: "When the channel was created",
              },
            },
            required: ["id", "displayName"],
          },
        },
        required: ["channel"],
      },
    },
  },
};
