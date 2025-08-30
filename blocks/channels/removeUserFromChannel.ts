import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const removeUserFromChannel: AppBlock = {
  name: "Remove Member from Channel",
  description:
    "Removes a member from a private Teams channel by their user ID. Note: Only works for private channels.\n\n**Required Permission:** ChannelMember.ReadWrite.All",
  category: "Channels",
  inputs: {
    default: {
      name: "Remove",
      description: "Trigger removing a member from the specified channel.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team containing the channel.",
          type: "string",
          required: true,
        },
        channelId: {
          name: "Channel ID",
          description:
            "The ID of the private channel to remove the member from.",
          type: "string",
          required: true,
        },
        userId: {
          name: "User ID",
          description: "The ID of the user to remove from the channel.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId, userId } = input.event.inputConfig;

        // First, get the membership ID for this user
        const membersResponse = await callGraphApi(
          `/teams/${teamId}/channels/${channelId}/members?$filter=userId eq '${userId}'`,
          accessToken,
        );

        const members = membersResponse.value || [];
        if (members.length === 0) {
          throw new Error(`User ${userId} is not a member of this channel`);
        }

        const membershipId = members[0].id;

        // Now remove the member using the membership ID
        await callGraphApi(
          `/teams/${teamId}/channels/${channelId}/members/${membershipId}`,
          accessToken,
          {
            method: "DELETE",
          },
        );

        await events.emit({
          teamId,
          channelId,
          userId,
          membershipId,
        });
      },
    },
  },
  outputs: {
    default: {
      name: "Member Removed",
      description:
        "Emitted when the member has been successfully removed from the channel.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          teamId: { type: "string", description: "Team ID" },
          channelId: { type: "string", description: "Channel ID" },
          userId: { type: "string", description: "User ID that was removed" },
          membershipId: {
            type: "string",
            description: "Membership ID that was deleted",
          },
        },
        required: ["teamId", "channelId", "userId", "membershipId"],
      },
    },
  },
};
