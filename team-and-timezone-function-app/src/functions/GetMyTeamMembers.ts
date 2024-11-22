import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphService } from "../services/graphService";
import { TeamMember } from "../types";
import { extractUpnFromToken } from "../utils/helpers/common";

export async function GetMyTeamMembers(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        let requestBody: any = {};
        let otherUserIds: string | null = null;
        let accessToken: string | null = null;

        if (req.method === 'POST') {
            requestBody = await req.json();
            otherUserIds = requestBody.otherUserIds || null;
            accessToken = requestBody.accessToken || null;
        }
        
        if (req.method === 'GET') {
            otherUserIds = req.query.get('otherUserIds') || null;
            accessToken = req.query.get('accessToken') || null;
        }

        // Get user impersonation access token from the request headers
        const userImpersonationAccessToken = req.headers.get("authorization")?.split(' ')[1];

        const graphService = new GraphService(accessToken, userImpersonationAccessToken);

        // Extract the "upn" claim from either token
        let upn: string | null = null;
        if (userImpersonationAccessToken) {
            upn = extractUpnFromToken(userImpersonationAccessToken);
        }
        
        if (!upn && accessToken) {
            upn = extractUpnFromToken(accessToken);
        }

        const teamMembers: TeamMember[] = await graphService.getTeamMembersDetails(otherUserIds);

        // Check if errors were encountered
        if (!teamMembers || teamMembers.length === 0) {
            return {
                status: 400,
                jsonBody: { message: "No valid team members data available, some requests may have failed." },
            };
        }

        // Sort the team members array
        if (upn) {
            teamMembers.sort((a, b) => {
                if (a.userPrincipalName === upn) return -1;
                if (b.userPrincipalName === upn) return 1;
                return 0;
            });
        }

        return { status: 200, jsonBody: teamMembers };

    } catch (error) {
        return { status: 500, jsonBody: { message: `Error fetching direct reports: ${error.message}` } };
    }
};

app.http('GetMyTeamMembers', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: GetMyTeamMembers
});
