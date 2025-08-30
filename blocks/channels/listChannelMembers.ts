import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const listChannelMembers: AppBlock = {
  name: "List Channel Members",
  description:
    "Lists all members of a Teams channel.\n\n**Required Permission:** ChannelMember.Read.All",
  category: "Channels",
  inputs: {
    default: {
      name: "List",
      description: "Trigger listing all members in the specified channel.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team containing the channel.",
          type: "string",
          required: true,
        },
        channelId: {
          name: "Channel ID",
          description: "The ID of the channel to list members from.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId } = input.event.inputConfig;

        const response = await callGraphApi(
          `/teams/${teamId}/channels/${channelId}/members`,
          accessToken,
        );

        await events.emit({ members: response.value || [] });
      },
    },
  },
  outputs: {
    default: {
      name: "Members Listed",
      description:
        "Emitted when channel members have been successfully retrieved.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          members: {
            type: "array",
            description: "Array of members in the channel",
            items: {
              type: "object",
              properties: {
                id: {
                  type: "string",
                  description: "The member ID",
                },
                displayName: {
                  type: "string",
                  description: "Member display name",
                },
                email: {
                  type: "string",
                  description: "Member email address",
                },
                roles: {
                  type: "array",
                  description: "Member roles (e.g., owner, member)",
                  items: { type: "string" },
                },
                userId: {
                  type: "string",
                  description: "The user ID",
                },
              },
              required: ["id", "displayName"],
            },
          },
        },
        required: ["members"],
      },
    },
  },
};
