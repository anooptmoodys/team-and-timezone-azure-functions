import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphService } from "../services/graphService";
import { TeamMember } from "../types";
import { validateToken } from "../utils/helpers/tokenHelper";
import { getErrorResponse } from "../utils/helpers/common";

// Helper to get required environment variables
const getEnv = (key: string, defaultValue: string = ""): string => process.env[key] || defaultValue;

// Extract parameters from the request (GET or POST)
const extractRequestParams = async (req: HttpRequest) => {
    if (req.method === "POST") {
        const body: any = await req.json();
        return {
            userId: body.userId || null,
            otherTeamMemberIds: body.otherTeamMemberIds || null,
            otherUsersOnly: body.otherUsersOnly || false
        };
    }
    return {
        userId: req.query.get("userId") || null,
        otherTeamMemberIds: req.query.get("otherTeamMemberIds") || null,
        otherUsersOnly: req.query.get("otherUsersOnly") || false
    };
};

// function to get team member by id from the team members list
const getTeamMemberById = (teamMembers: TeamMember[], id: string): TeamMember => {
    return teamMembers.find(member => member.id === id);
};

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
        const token = req.headers.get("Authorization")?.replace("Bearer ", "") || "";
        const tokenValidationResult = await validateToken(token);

        if (!tokenValidationResult.valid) {
            return getErrorResponse(401, "Unauthorized", `Token validation failed. Details: ${tokenValidationResult.errorMessage}`);
        }

        const { userId, otherTeamMemberIds, otherUsersOnly } = await extractRequestParams(req);

        const graphService = new GraphService(userId);

        const teamMembers = await graphService.getTeamMembersDetails(otherTeamMemberIds, otherUsersOnly);

        if (!teamMembers || teamMembers.length === 0) {
            return getErrorResponse(400, "BadRequest", "The request is invalid.");
        }

        const requestedByUser = getTeamMemberById(teamMembers, userId);

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
                        name: requestedByUser?.department || null,
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