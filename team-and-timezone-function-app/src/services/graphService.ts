import { User } from "@microsoft/microsoft-graph-types";
import { TeamMember } from "../types";
import { createAppGraphClient } from "../utils/graphClient/appGraphClient";
import { createDelegatedGraphClient, createDelegatedGraphClientWithOnBehalfOf } from "../utils/graphClient/delegatedGraphClient";
import { Client } from "@microsoft/microsoft-graph-client";
import { getIanaFromWindows } from "../utils/helpers/common";

export class GraphService {
    private delegatedClient: Client;
    private appClient: Client;
    private propertiesToSelectForUser = ["id", "displayName", "mail", "userPrincipalName", "officeLocation", "city", "country", "jobTitle"];

    constructor(private accessToken?: string, private userAccessToken?: string) {

        // either userAccessToken or accessToken is required
        if (!this.accessToken && !this.userAccessToken) {
            throw new Error("Access token is required for delegated client.");
        }

        if (this.userAccessToken) {
            this.delegatedClient = createDelegatedGraphClientWithOnBehalfOf(this.userAccessToken);
        }

        if (this.accessToken) {
            this.delegatedClient = createDelegatedGraphClient(this.accessToken);
        }

        const clientId = process.env.CLIENT_ID;
        const tenantId = process.env.TENANT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        this.appClient = createAppGraphClient(clientId, tenantId, clientSecret);
    }

    /**
     * Fetch the user and user's direct reports.
     */
    private async getUserAndDirectReports(): Promise<User[]> {
        try {
            // First, get your own user details by querying "/me"
            const meRequest = {
                id: "me",
                method: "GET",
                url: "/me?$select=" + this.propertiesToSelectForUser.join(","),
            };

            // Then, get the direct reports
            const directReportsRequest = {
                id: "directReports",
                method: "GET",
                url: "/me/directReports?$select=" + this.propertiesToSelectForUser.join(","),
            };

            // Combine both requests into a batch request
            const batchRequest = {
                requests: [meRequest, directReportsRequest],
            };

            // Make the batch request to fetch both /me and direct reports
            const batchResponse = await this.delegatedClient.api("/$batch").post(batchRequest);

            // Extract the responses for /me and /me/directReports
            const meResponse = batchResponse.responses.find((r: any) => r.id === "me");
            const directReportsResponse = batchResponse.responses.find((r: any) => r.id === "directReports");

            // Combine the results
            const directReports = directReportsResponse?.body?.value || [];
            const meData = meResponse?.body;

            // Include yourself in the team members list (add yourself to the direct reports list)
            if (meData) {
                directReports.unshift(meData);
            }

            return directReports;
        } catch (error) {
            console.error("Error fetching direct reports and user details:", error);
            return [];
        }
    }

    /**
     * Fetch my manager's details, and my manager's direct reports.
     */
    private async getManagerAndDirectReports(): Promise<User[]> {
        try {
            // Step 1: Fetch the manager's details
            const managerResponse = await this.delegatedClient
                .api("/me/manager")
                .select(this.propertiesToSelectForUser)
                .get();

            const managerId = managerResponse?.id;

            // If no manager is found, return an empty array
            if (!managerId) {
                console.warn("No manager found for the current user.");
                return [];
            }

            // Step 2: Fetch the manager's direct reports using the managerId
            const directReportsResponse = await this.delegatedClient
                .api(`/users/${managerId}/directReports`)
                .select(this.propertiesToSelectForUser)
                .get();

            const managerDirectReports = directReportsResponse?.value || [];

            // Combine the manager's details with their direct reports
            return [managerResponse, ...managerDirectReports];
        } catch (error) {
            console.error("Error fetching manager and manager's direct reports:", error);
            return [];
        }
    }


    private async fetchDirectReportsOrManagerReports(): Promise<User[]> {
        let teamMembers = await this.getUserAndDirectReports();

        // If teamMembers has only my details, fetch my manager's details and their direct reports
        if (!teamMembers || teamMembers.length === 1) {
            const managerAndDirectReports = await this.getManagerAndDirectReports();
            teamMembers = managerAndDirectReports && managerAndDirectReports.length > 0 ? managerAndDirectReports : teamMembers;
        }

        return teamMembers;
    }

    private async fetchCustomUserDetails(customUserIds: string[]): Promise<User[]> {

        // if no customUserIds are provided, return empty array
        if (customUserIds.length === 0) {
            return [];
        }

        const delegatedBatchRequests = customUserIds.map((id) => ({
            id: `user|${id}`,
            method: "GET",
            url: `/users/${id}?$select=${this.propertiesToSelectForUser.join(",")}`,
        }));

        const delegatedBatchResponse = await this.delegatedClient.api("/$batch").post({ requests: delegatedBatchRequests });

        return delegatedBatchResponse.responses
            .filter((r: any) => r.id.startsWith("user|"))
            .map((r: any) => r.body);
    }

    private async fetchPresenceData(allUserIds: string[]): Promise<any> {

        // if no user ids are provided, return empty object
        if (allUserIds.length === 0) {
            return {};
        }

        const delegatedBatchRequests = [{
            id: "presence",
            method: "POST",
            headers: { "Content-Type": "application/json" },
            url: `/communications/getPresencesByUserId`,
            body: { ids: allUserIds },
        }];

        const delegatedBatchResponse = await this.delegatedClient.api("/$batch").post({ requests: delegatedBatchRequests });

        const presenceResponse = delegatedBatchResponse.responses.find((r: any) => r.id === "presence");
        return presenceResponse?.body?.value.reduce((acc: any, item: any) => {
            acc[item.id] = item.availability;
            return acc;
        }, {});
    }

    private async fetchTimezoneData(allUserIds: string[]): Promise<any> {

        // if no user ids are provided, return empty object
        if (allUserIds.length === 0) {
            return {};
        }

        const appBatchRequests = allUserIds.map((id) => ({
            id: `timezone|${id}`,
            method: "GET",
            url: `/users/${id}/mailboxSettings/timeZone`,
        }));

        const appBatchResponse = await this.appClient.api("/$batch").post({ requests: appBatchRequests });

        return appBatchResponse.responses.reduce((acc: any, response: any) => {
            if (response.id.startsWith("timezone|")) {
                const userId = response.id.split("|")[1];
                acc[userId] = response?.body?.value || null;
            }
            return acc;
        }, {});
    }

    async getTeamMembersDetails(otherUserIds?: string): Promise<TeamMember[]> {
        const teamMembers = await this.fetchDirectReportsOrManagerReports();
        let customUserIds = otherUserIds ? otherUserIds.split(";").map((id) => id.trim()) : [];
        // if any Id is repeated in both teamMembers and customUserIds, remove it from customUserIds
        customUserIds = customUserIds.filter((id) => !teamMembers.some((user) => user.id === id));
        const allUsers = [...teamMembers];
        const allUserIds = Array.from(new Set([...allUsers.map((user) => user.id), ...customUserIds]));

        console.log(allUserIds)

        const [customUsers, presenceData, timezoneData] = await Promise.all([
            this.fetchCustomUserDetails(customUserIds),
            this.fetchPresenceData(allUserIds),
            this.fetchTimezoneData(allUserIds)
        ]);

        const mapUserDetails = (user: User): TeamMember => {
            // location should be officeLocation, city, country
            const location = [user.officeLocation, user.city, user.country].filter(Boolean).join(", ");
            return {
                id: user.id,
                name: user.displayName,
                mail: user.mail || null,
                userPrincipalName: user.userPrincipalName || null,
                location: location || null,
                jobTitle: user.jobTitle || null,
                presence: presenceData[user.id] || null,
                timeZone: getIanaFromWindows(timezoneData[user.id] ?? "GMT Standard Time"),
                photo: `/_layouts/15/userphoto.aspx?size=L&username=${user.userPrincipalName}`,
            };
        }

        const customUsersDetails = customUsers.map((user: any) => (mapUserDetails(user)));

        const teamMemberDetails = allUsers.map((user) => (mapUserDetails(user)));

        return [...teamMemberDetails, ...customUsersDetails];
    }

    // function to get the users presence based on the users ids passed a string separated by ;
    async getUsersPresence(userIds: string): Promise<any> {
        const allUserIds = userIds.split(";").map((id) => id.trim());
        const presenceData = await this.fetchPresenceData(allUserIds);
        return presenceData;
    }

}
