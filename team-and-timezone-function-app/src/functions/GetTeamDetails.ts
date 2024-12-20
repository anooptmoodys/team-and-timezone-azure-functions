import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphService } from "../services/graphService";
import { validateToken } from "../utils/helpers/tokenHelper";
import { getErrorResponse, getMockData, extractRequestParams } from "../utils/helpers/common";

// Sample return data

/* 
//no error
{
    "status": 200,
    "error": {
        "exists": false,
        "code": null,
        "message": null
    },
    "data": {
        "team": {
            "name": "Team A",
            "members": [{"id":"02485fea-5da8-4277-9da5-d1147470196a","name":"Anoop Tatti","mail":"tattia@mcolab.moodys.com","userPrincipalName":"tattia@mcolab.moodys.com","location":null,"jobTitle":null,"presence":null,"timeZone":"Europe/London","photo":"/_layouts/15/userphoto.aspx?size=L&username=tattia@mcolab.moodys.com"}]
        }
    }
}

//error
{
    "status": 400,
    "error": {
        "exists": true,
        "code": "BadRequest",
        "message": "The request is invalid."
    },
    "data": null
}

*/

export async function GetTeamDetails(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {

        // Validate token - by getting the token from the request headers and removing the "Bearer " prefix
        /* const token = req.headers.get("Authorization")?.replace("Bearer ", "") || "";
        const tokenValidationResult = await validateToken(token);

        if (!tokenValidationResult.valid) {
            return getErrorResponse(401, "Unauthorized", `Token validation failed. Details: ${tokenValidationResult.errorMessage}`);
        } */

        const { userId, teamMembersIds } = await extractRequestParams(req, ["userId", "teamMembersIds"]);

        const graphService = new GraphService(userId);

        // const teamMembers = teamMembersIds ? await graphService.getTeamMembersDetailsByIds(teamMembersIds) : await graphService.getTeamMembersDetails();
        const teamMembers = teamMembersIds ? await graphService.getTeamMembersDetailsByIds(teamMembersIds) : getMockData();

        if (!teamMembers || teamMembers.length === 0) {
            return getErrorResponse(400, "BadRequest", "The request is invalid.");
        }
        
        return { 
            status: 200, 
            jsonBody: {
                status: 200,
                error: {
                    exists: false,
                    code: null,
                    message: null
                },
                data: {
                    team: {
                        name: null,
                        members: teamMembers
                    }
                }
            }
        }
    } catch (error) {
        return getErrorResponse(500, "InternalServerError", error.message);
    }
}

app.http("GetTeamDetails", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: GetTeamDetails,
});