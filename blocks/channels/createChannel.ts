import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const createChannel: AppBlock = {
  name: "Create Channel",
  description:
    "Creates a new channel in a Teams team.\n\n**Required Permission:** Channel.ReadWrite.All",
  category: "Channels",
  inputs: {
    default: {
      name: "Create",
      description: "Trigger creating a new channel.",
      config: {
        teamId: {
          name: "Team ID",
          description: "The ID of the team to create the channel in.",
          type: "string",
          required: true,
        },
        displayName: {
          name: "Channel Name",
          description: "Name of the channel to create.",
          type: "string",
          required: true,
        },
        description: {
          name: "Description",
          description: "Description of the channel (optional).",
          type: "string",
          required: false,
        },
        membershipType: {
          name: "Membership Type",
          description:
            "Type of channel: 'standard' (default) or 'private'. Private channels require owner membership to be specified.",
          type: "string",
          required: false,
          default: "standard",
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { teamId, displayName, description, membershipType } =
          input.event.inputConfig;

        const channelPayload: any = {
          displayName,
          membershipType: membershipType || "standard",
        };

        if (description) {
          channelPayload.description = description;
        }

        const channel = await callGraphApi(
          `/teams/${teamId}/channels`,
          accessToken,
          {
            method: "POST",
            body: channelPayload,
          },
        );

        await events.emit({
          channel,
        });
      },
    },
  },
  outputs: {
    default: {
      name: "Channel Created",
      description: "Emitted when the channel has been successfully created.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          channel: {
            type: "object",
            description: "The created channel object.",
            properties: {
              id: { type: "string" },
              displayName: { type: "string" },
              description: { type: "string" },
              membershipType: { type: "string" },
              webUrl: { type: "string" },
            },
            required: ["id", "displayName"],
          },
        },
        required: ["channel"],
      },
    },
  },
};
