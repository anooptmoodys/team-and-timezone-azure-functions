import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphService } from "../services/graphService";

// Extract parameters from the request (GET or POST)
const extractRequestParams = async (req: HttpRequest) => {
    if (req.method === "POST") {
        const body: any = await req.json();
        return {
            userIds: body.userIds || null
        };
    }
    return {
        userIds: req.query.get("userIds") || null
    };
};

export async function GetUsersPresence(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const { userIds } = await extractRequestParams(req);

        const graphService = new GraphService();

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
