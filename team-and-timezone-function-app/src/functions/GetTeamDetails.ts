import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphService } from "../services/graphService";
import { TeamMember } from "../types";
import { extractUpnFromToken } from "../utils/helpers/common";

// Helper to get required environment variables
const getEnv = (key: string, defaultValue: string = ""): string => process.env[key] || defaultValue;

// Extract parameters from the request (GET or POST)
const extractRequestParams = async (req: HttpRequest) => {
    if (req.method === "POST") {
        const body: any = await req.json();
        return {
            userId: body.userId || null,
            otherUserIds: body.otherUserIds || null
        };
    }
    return {
        userId: req.query.get("userId") || null,
        otherUserIds: req.query.get("otherUserIds") || null
    };
};

// Sort team members, prioritizing the current user's id
const sortTeamMembers = (teamMembers: TeamMember[], id: string | null): TeamMember[] => {
    if (!id) return teamMembers;
    return teamMembers.sort((a, b) => {
        if (a.id === id) return -1;
        if (b.id === id) return 1;
        return 0;
    });
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
        const { userId, otherUserIds } = await extractRequestParams(req);

        const graphService = new GraphService(userId);

        const teamMembers = await graphService.getTeamMembersDetails(otherUserIds);

        if (!teamMembers || teamMembers.length === 0) {
            return {
                status: 400,
                jsonBody: {
                    status: 400,
                    error: {
                        exists: true,
                        code: "BadRequest",
                        message: "The request is invalid."
                    },
                    data: null
                }
            };
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
                        name: "Team A",
                        members: teamMembers
                    }
                }
            }
        }
    } catch (error) {
        return { 
            status: 500,
            jsonBody: {
                status: 500,
                error: {
                    exists: true,
                    code: "InternalServerError",
                    message: error.message
                },
                data: null
            }
        };
    }
}

app.http("GetTeamDetails", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: GetTeamDetails,
});