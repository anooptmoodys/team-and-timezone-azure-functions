import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphService } from "../services/graphService";
import { getErrorResponse, extractRequestParams } from "../utils/helpers/common";
import { validateToken } from "../utils/helpers/tokenHelper";


export async function GetUsersPresence(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {

         // Validate token - by getting the token from the request headers and removing the "Bearer " prefix
        /* const token = req.headers.get("Authorization")?.replace("Bearer ", "") || "";
        const tokenValidationResult = await validateToken(token);

        if (!tokenValidationResult.valid) {
            return getErrorResponse(401, "Unauthorized", `Token validation failed. Details: ${tokenValidationResult.errorMessage}`);
        } */
       
        const { teamMembersIds } = await extractRequestParams(req, ["teamMembersIds"]);

        const graphService = new GraphService();

        // Get the presence data for the provided user ids
        const presenceData = await graphService.getUsersPresence(teamMembersIds);

        // Check if errors were encountered
        if (!presenceData || presenceData.length === 0) {
            return getErrorResponse(400, "BadRequest", "No valid presence data available, some requests may have failed.");
        }

        return { status: 200, jsonBody: presenceData };

    }
    catch (error) {
        return getErrorResponse(500, "InternalServerError", `Error fetching presence data: ${error.message}`);
    }
};

app.http('GetUsersPresence', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: GetUsersPresence
});
