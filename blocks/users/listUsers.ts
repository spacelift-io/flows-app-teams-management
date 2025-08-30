import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";
import { userSchema } from "../../schemas/userSchema.ts";

export const listUsers: AppBlock = {
  name: "List Users",
  description:
    "Lists users in the organization. Supports filtering and searching by various properties.\n\n**Required Permission:** User.Read.All",
  category: "Users",
  inputs: {
    default: {
      name: "List Users",
      description: "Retrieve a list of users.",
      config: {
        filter: {
          name: "Filter",
          description:
            "OData filter expression (e.g., startsWith(displayName,'John'), endsWith(mail,'.edu'))",
          type: "string",
          required: false,
        },
        search: {
          name: "Search",
          description:
            "Search across user properties (e.g., displayName:John). Requires advanced query.",
          type: "string",
          required: false,
        },
        top: {
          name: "Max Results",
          description: "Maximum number of users to return (default 100)",
          type: "number",
          required: false,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const { filter, search, top } = input.event.inputConfig;

        const params = new URLSearchParams();
        params.append(
          "$select",
          "id,userPrincipalName,displayName,givenName,surname,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone,accountEnabled",
        );
        if (filter) params.append("$filter", filter);
        if (search) params.append("$search", `"${search}"`);
        if (top) params.append("$top", top.toString());

        const queryString = params.toString();
        const endpoint = `/users?${queryString}`;

        const response = await callGraphApi(endpoint, accessToken, {
          method: "GET",
        });

        const users = response.value || [];
        await events.emit({ users });
      },
    },
  },
  outputs: {
    default: {
      name: "Users",
      description: "List of users matching the query",
      type: {
        type: "object",
        properties: {
          users: {
            type: "array",
            description: "Array of user objects",
            items: userSchema,
          },
        },
        required: ["users"],
      },
    },
  },
};
