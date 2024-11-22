import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphService } from "../services/graphService";

export async function GetUsersPresence(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        let requestBody: any = {};
        let userIds: string | null = null;
        let accessToken: string | null = null;

        if (request.method === 'POST') {
            requestBody = await request.json();
            userIds = requestBody.userIds || null;
            accessToken = requestBody.accessToken || null;
        }

        if (request.method === 'GET') {
            userIds = request.query.get('userIds') || null;
            accessToken = request.query.get('accessToken') || null;
        }

        // Get user impersonation access token from the request headers
        const userImpersonationAccessToken = request.headers.get("authorization")?.split(' ')[1];

        const graphService = new GraphService(accessToken, userImpersonationAccessToken);

        // Get the presence data for the provided user ids
        const presenceData = await graphService.getUsersPresence(userIds);

        // Check if errors were encountered
        if (!presenceData || presenceData.length === 0) {
            return {
                status: 400,
                jsonBody: { message: "No valid presence data available, some requests may have failed." },
            };
        }

        return { status: 200, jsonBody: presenceData };

    }
    catch (error) {
        return { status: 500, jsonBody: { message: `Error fetching presence data: ${error.message}` } };
    }
};

app.http('GetUsersPresence', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: GetUsersPresence
});
