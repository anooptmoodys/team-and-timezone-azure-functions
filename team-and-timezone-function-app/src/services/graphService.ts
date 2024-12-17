import { User } from "@microsoft/microsoft-graph-types";
import { TeamMember } from "../types";
import { createAppGraphClient } from "../utils/graphClient/appGraphClient";
import { Client } from "@microsoft/microsoft-graph-client";
import { getIanaFromWindows } from "../utils/helpers/common";
import { mapTeamMemberDetails } from "../utils/helpers/mappers";

export class GraphService {
    private appClient: Client;
    private corePropertiesToSelectForUser = ["id", "givenName", "surname", "displayName", "userPrincipalName", "officeLocation", "jobTitle", "department"];
    private propertiesToSelectForUser = [...this.corePropertiesToSelectForUser, "mail", "city", "country"];

    constructor(private userId?: string) {

        this.userId = userId;
        const clientId = process.env.CLIENT_ID;
        const tenantId = process.env.TENANT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        this.appClient = createAppGraphClient(clientId, tenantId, clientSecret);
    }

    private async getUserAndDirectReports(): Promise<User[]> {
        try {
            // First, get user details by querying "/users/${this.userId}"
            const userRequest = {
                id: "user",
                method: "GET",
                url: `/users/${this.userId}?$select=${this.propertiesToSelectForUser.join(",")}`,
            };

            // Then, get the direct reports
            const directReportsRequest = {
                id: "directReports",
                method: "GET",
                url: `/users/${this.userId}/directReports?$select=${this.propertiesToSelectForUser.join(",")}`,
            };

            // Combine both requests into a batch request
            const batchRequest = {
                requests: [userRequest, directReportsRequest],
            };

            // Make the batch request to fetch both user and direct reports
            const batchResponse = await this.appClient.api("/$batch").post(batchRequest);

            // Extract the responses for /user and /user/directReports
            const userResponse = batchResponse.responses.find((r: any) => r.id === "user");
            const directReportsResponse = batchResponse.responses.find((r: any) => r.id === "directReports");

            // Combine the results
            const directReports = directReportsResponse?.body?.value || [];
            const userData = userResponse?.body;

            // Include yourself in the team members list (add yourself to the direct reports list)
            if (userData) {
                directReports.unshift(userData);
            }

            return directReports;
        } catch (error) {
            console.error("Error fetching direct reports and user details:", error);
            return [];
        }
    }

    /**
     * Fetch user's manager's details, and user's manager's direct reports.
     */
    private async getManagerAndDirectReports(): Promise<User[]> {
        try {
            // Step 1: Fetch the manager's details
            const managerResponse = await this.appClient
                .api(`/users/${this.userId}/manager`)
                .select(this.propertiesToSelectForUser)
                .get();

            const managerId = managerResponse?.id;

            // If no manager is found, return an empty array
            if (!managerId) {
                console.warn("No manager found for the current user.");
                return [];
            }

            // Step 2: Fetch the manager's direct reports using the managerId
            const directReportsResponse = await this.appClient
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

    /**
     * Fetch people user works with
     */
    private async getPeopleWorkingWith(): Promise<User[]> {
        try {
            const peopleWorkingWithResponse = await this.appClient
                .api(`/users/${this.userId}/people`)
                .select(["id"])
                .get();

            return peopleWorkingWithResponse?.value || [];
        } catch (error) {
            console.error("Error fetching people user works with:", error);
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

    private async fetchUsersByIds(userIds: string[]): Promise<any[]> {

        // if no userIds are provided, return empty array
        if (userIds.length === 0) {
            return [];
        }

        const appBatchRequests = userIds.map((id) => ({
            id: `user|${id}`,
            method: "GET",
            url: `/users/${id}?$select=${this.propertiesToSelectForUser.join(",")}`,
        }));

        const appBatchResponse = await this.appClient.api("/$batch").post({ requests: appBatchRequests });

        return appBatchResponse.responses
            .filter((r: any) => r.id.startsWith("user|"))
            .map((r: any) => r.body);
    }

    private async fetchPresenceData(allUserIds: string[]): Promise<any> {

        // if no user ids are provided, return empty object
        if (allUserIds.length === 0) {
            return {};
        }

        const appBatchRequests = [{
            id: "presence",
            method: "POST",
            headers: { "Content-Type": "application/json" },
            url: `/communications/getPresencesByUserId`,
            body: { ids: allUserIds },
        }];

        const appBatchResponse = await this.appClient.api("/$batch").post({ requests: appBatchRequests });

        // if appBatchResponse.status is not 200, return empty object
        if (appBatchResponse.status !== 200) {
            return {};
        }

        const presenceResponse = appBatchResponse.responses.find((r: any) => r.id === "presence");

        return presenceResponse?.body?.value?.reduce((acc: any, item: any) => {
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

        return appBatchResponse?.responses?.reduce((acc: any, response: any) => {
            if (response.id.startsWith("timezone|")) {
                const userId = response.id.split("|")[1];
                acc[userId] = response?.body?.value || null;
            }
            return acc;
        }, {});
    }

    // function to get team members details by Ids
    async getTeamMembersDetailsByIds(teamMemberIds: string): Promise<TeamMember[]> {
        const allTeamMemberIds = teamMemberIds.split(";").map((id) => id.trim());
        const allTeamMembers = await this.fetchUsersByIds(allTeamMemberIds);

        const [presenceData, timezoneData] = await Promise.all([
            this.fetchPresenceData(allTeamMemberIds),
            this.fetchTimezoneData(allTeamMemberIds)
        ]);

        const teamMemberDetails: TeamMember[] = allTeamMembers.map((teamMember) => (mapTeamMemberDetails(teamMember, presenceData, timezoneData)));

        return teamMemberDetails;
    }

    // function to get core team members details and people user works with
    async getTeamMembersDetails(): Promise<TeamMember[]> {
        const coreTeamMembers = await this.fetchDirectReportsOrManagerReports();
        const peopleWorkingWith = await this.getPeopleWorkingWith();
        const peopleWorkingWithIds = peopleWorkingWith.map((user) => user.id);
        const nonCoreTeamMembers = await this.fetchUsersByIds(peopleWorkingWithIds);

        const allUsers = [...coreTeamMembers, ...nonCoreTeamMembers];
        const allUserIds = allUsers.map((user) => user.id);

        const [presenceData, timezoneData] = await Promise.all([
            this.fetchPresenceData(allUserIds),
            this.fetchTimezoneData(allUserIds)
        ]);

        const mappedTeamMembersDetails = allUsers.map((user) => (mapTeamMemberDetails(user, presenceData, timezoneData)));

        // add isCoreTeamMember property to the team members
        const teamMembersDetails: TeamMember[] = mappedTeamMembersDetails.map((teamMember) => ({
            ...teamMember,
            isCoreTeamMember: coreTeamMembers.some((u) => u.id === teamMember.id)
        }));

        return teamMembersDetails;
    }

    /* async getTeamMembersDetails(otherTeamMemberIds?: string, otherUsersOnly?: boolean): Promise<TeamMember[]> {
        const teamMembers = otherUsersOnly ? [] : await this.fetchDirectReportsOrManagerReports();
        let customUserIds = otherTeamMemberIds ? otherTeamMemberIds.split(";").map((id) => id.trim()) : [];
        const customUsers = await this.fetchUsersByIds(customUserIds);

        let allUsers: User[] = [...teamMembers, ...customUsers];
        // make sure allUsers is unique
        allUsers = allUsers.filter((user, index, self) => self.findIndex((u) => u.id === user.id) === index);
        const allUserIds = allUsers.map((user) => user.id);

        const [presenceData, timezoneData] = await Promise.all([
            this.fetchPresenceData(allUserIds),
            this.fetchTimezoneData(allUserIds)
        ]);

        const mappedTeamMembersDetails = allUsers.map((user) => (mapTeamMemberDetails(user, presenceData, timezoneData)));

        // add isOtherTeamMember property to the team members
        const teamMembersDetails: TeamMember[] = mappedTeamMembersDetails.map((teamMember) => ({
            ...teamMember,
            isOtherTeamMember: customUsers.some((u) => u.id === teamMember.id)
        }));

        return teamMembersDetails;
    } */

    // function to get the users presence based on the users ids passed a string separated by ;
    async getUsersPresence(userIds: string): Promise<any> {
        const allUserIds = userIds.split(";").map((id) => id.trim());
        const presenceData = await this.fetchPresenceData(allUserIds);
        return presenceData;
    }
}