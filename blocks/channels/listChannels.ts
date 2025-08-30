import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const listChannels: AppBlock = {
  name: "List Channels",
  description:
    "Lists all channels in a Team/Community.\n\n**Required Permission:** Channel.ReadBasic.All",
  category: "Channels",
  inputs: {
    default: {
      name: "List",
      description: "Trigger listing all channels in the specified team.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team to list channels from.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId } = input.event.inputConfig;

        const response = await callGraphApi(
          `/teams/${teamId}/channels`,
          accessToken,
        );

        await events.emit({ channels: response.value || [] });
      },
    },
  },
  outputs: {
    default: {
      name: "Channels Listed",
      description:
        "Emitted when channels have been successfully retrieved from the team.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          channels: {
            type: "array",
            description: "Array of channels in the team",
            items: {
              type: "object",
              properties: {
                id: {
                  type: "string",
                  description: "The Channel ID (use this in other blocks)",
                },
                displayName: { type: "string", description: "Channel name" },
                description: {
                  type: "string",
                  description: "Channel description",
                },
                email: {
                  type: "string",
                  description: "Email address for the channel",
                },
                webUrl: {
                  type: "string",
                  description: "Web URL to the channel",
                },
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
        },
        required: ["channels"],
      },
    },
  },
};
