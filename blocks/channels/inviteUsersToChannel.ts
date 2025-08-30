import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const inviteUsersToChannel: AppBlock = {
  name: "Add Members to Channel",
  description:
    "Adds members to a private Teams channel. Note: Only works for private channels.\n\n**Required Permission:** ChannelMember.ReadWrite.All",
  category: "Channels",
  inputs: {
    default: {
      name: "Add",
      description: "Trigger adding members to the specified channel.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team containing the channel.",
          type: "string",
          required: true,
        },
        channelId: {
          name: "Channel ID",
          description: "The ID of the private channel to add members to.",
          type: "string",
          required: true,
        },
        userIds: {
          name: "User IDs",
          description: "Array of user IDs to add as members.",
          type: {
            type: "array",
            items: { type: "string" },
          },
          required: true,
        },
        roles: {
          name: "Roles",
          description:
            'Member roles. Can be empty array for standard member, or ["owner"] to make them an owner.',
          type: {
            type: "array",
            items: { type: "string" },
          },
          required: false,
          default: [],
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId, userIds, roles } = input.event.inputConfig;

        const addedMembers = [];

        for (const userId of userIds) {
          const memberPayload = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: roles || [],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`,
          };

          const member = await callGraphApi(
            `/teams/${teamId}/channels/${channelId}/members`,
            accessToken,
            {
              method: "POST",
              body: memberPayload,
            },
          );

          addedMembers.push(member);
        }

        await events.emit({
          teamId,
          channelId,
          members: addedMembers,
        });
      },
    },
  },
  outputs: {
    default: {
      name: "Members Added",
      description:
        "Emitted when members have been successfully added to the channel.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          teamId: { type: "string" },
          channelId: { type: "string" },
          members: {
            type: "array",
            items: {
              type: "object",
              properties: {
                id: { type: "string" },
                displayName: { type: "string" },
                roles: {
                  type: "array",
                  items: { type: "string" },
                },
              },
            },
          },
        },
        required: ["teamId", "channelId", "members"],
      },
    },
  },
};
