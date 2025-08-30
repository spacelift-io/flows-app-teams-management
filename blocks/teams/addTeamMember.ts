import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const addTeamMember: AppBlock = {
  name: "Add Team Member",
  description:
    "Adds a member to a Team.\n\n**Required Permission:** GroupMember.ReadWrite.All",
  category: "Teams",
  inputs: {
    default: {
      name: "Add",
      description: "Trigger adding a member to the specified team.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team to add the member to.",
          type: "string",
          required: true,
        },
        userId: {
          name: "User ID",
          description: "The ID of the user to add as a team member.",
          type: "string",
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
        const { teamId, userId, roles } = input.event.inputConfig;

        const memberPayload = {
          "@odata.type": "#microsoft.graph.aadUserConversationMember",
          roles: roles || [],
          "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`,
        };

        const member = await callGraphApi(
          `/teams/${teamId}/members`,
          accessToken,
          {
            method: "POST",
            body: memberPayload,
          },
        );

        await events.emit({
          teamId,
          member,
        });
      },
    },
  },
  outputs: {
    default: {
      name: "Member Added",
      description:
        "Emitted when member has been successfully added to the team.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          teamId: { type: "string" },
          member: {
            type: "object",
            properties: {
              id: { type: "string", description: "Member ID" },
              displayName: {
                type: "string",
                description: "Member display name",
              },
              userId: { type: "string", description: "User ID" },
              email: { type: "string", description: "Email address" },
              roles: {
                type: "array",
                items: { type: "string" },
                description: "Member roles",
              },
            },
          },
        },
        required: ["teamId", "member"],
      },
    },
  },
};
