import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";

export const listTeams: AppBlock = {
  name: "List Teams",
  description:
    "Lists all Teams available to the app with their IDs and names.\n\n**Required Permission:** Group.Read.All",
  category: "Teams",
  inputs: {
    default: {
      name: "List",
      description: "Trigger listing all teams available to the app.",
      config: {},
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);

        // Get all groups that are team-enabled
        const response = await callGraphApi(
          "/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description,visibility,createdDateTime",
          accessToken,
        );

        const teams = response.value || [];

        await events.emit({ teams });
      },
    },
  },
  outputs: {
    default: {
      name: "Teams Listed",
      description: "Emitted when teams have been successfully retrieved.",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          teams: {
            type: "array",
            description: "Array of teams available to the app",
            items: {
              type: "object",
              properties: {
                id: {
                  type: "string",
                  description: "The Team ID (use this in other blocks)",
                },
                displayName: { type: "string", description: "Team name" },
                description: {
                  type: "string",
                  description: "Team description",
                },
                visibility: {
                  type: "string",
                  description: "Public, Private, or HiddenMembership",
                },
                createdDateTime: {
                  type: "string",
                  description: "When the team was created",
                },
              },
              required: ["id", "displayName"],
            },
          },
        },
        required: ["teams"],
      },
    },
  },
};
