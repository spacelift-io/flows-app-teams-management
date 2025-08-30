import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const unarchiveChannel: AppBlock = {
  name: "Unarchive Channel",
  description:
    "Unarchives a previously archived Teams channel.\n\n**Required Permission:** Channel.ReadWrite.All",
  category: "Channels",
  inputs: {
    default: {
      name: "Unarchive",
      description: "Trigger unarchiving the specified channel.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team containing the channel.",
          type: "string",
          required: true,
        },
        channelId: {
          name: "Channel ID",
          description: "The ID of the channel to unarchive.",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, channelId } = input.event.inputConfig;

        await callGraphApi(
          `/teams/${teamId}/channels/${channelId}/unarchive`,
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
      name: "Channel Unarchived",
      description: "Emitted when the channel has been successfully unarchived.",
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
