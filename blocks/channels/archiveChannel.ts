import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const archiveChannel: AppBlock = {
  name: "Archive Channel",
  description:
    "Archives a Teams channel. Archived channels are read-only but can be unarchived later.\n\n**Required Permission:** Channel.ReadWrite.All",
  category: "Channels",
  inputs: {
    default: {
      name: "Archive",
      description: "Trigger archiving the specified channel.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team containing the channel.",
          type: "string",
          required: true,
        },
        channelId: {
          name: "Channel ID",
          description: "The ID of the channel to archive.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId } = input.event.inputConfig;

        await callGraphApi(
          `/teams/${teamId}/channels/${channelId}/archive`,
          accessToken,
          {
            method: "POST",
          },
        );

        await events.emit({
          teamId,
          channelId,
        });
      },
    },
  },
  outputs: {
    default: {
      name: "Channel Archived",
      description: "Emitted when the channel has been successfully archived.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          teamId: { type: "string" },
          channelId: { type: "string" },
        },
        required: ["teamId", "channelId"],
      },
    },
  },
};
