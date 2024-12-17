import { findIana } from "windows-iana";

// function to get the IANA time zone from the Windows time zone
export function getIanaFromWindows(windowsTimeZone: string): string {
    const ianaZones = findIana(windowsTimeZone);
    if (!ianaZones || ianaZones.length === 0) {
        return "Europe/London";
    }
    return ianaZones[0];
}

// function to return error response
export function getErrorResponse(status: number, code: string, message: string): any {
    return {
        status,
        jsonBody: {
            status,
            error: {
                exists: true,
                code,
                message
            },
            data: null
        }
    };
}

// function get mock data
export function getMockData(): any {
    return [
        {
            "id": "02485fea-5da8-4277-9da5-d1147470196a",
            "name": "Anoop Tatti",
            "mail": "tattia@mcolab.moodys.com",
            "userPrincipalName": "tattia@mcolab.moodys.com",
            "city": null,
            "country": null,
            "jobTitle": null,
            "presence": null,
            "timeZone": "Europe/London",
            "department": null,
            "photo": "/_layouts/15/userphoto.aspx?size=L&username=tattia@mcolab.moodys.com",
            "isCoreTeamMember": true
        },
        {
            "id": "a59b53d8-259d-4098-80f8-350398477cab",
            "name": "Navneet Singh",
            "mail": "Navneet.Singh@mcolab.moodys.com",
            "userPrincipalName": "singhnav@mcolab.moodys.com",
            "city": "Mumbai",
            "country": "India",
            "jobTitle": "Non-Payroll Contractor",
            "presence": null,
            "timeZone": "Asia/Calcutta",
            "department": "TSG",
            "photo": "/_layouts/15/userphoto.aspx?size=L&username=singhnav@mcolab.moodys.com",
            "isCoreTeamMember": true
        },
        {
            "id": "a84118fa-9a0b-444f-8bd1-08c99d2ada0a",
            "name": "Dikshant Dwivedi",
            "mail": "Dikshant.Dwivedi@mcolab.onmicrosoft.com",
            "userPrincipalName": "DwivediDi@mcolab.moodys.com",
            "city": "Gurugram",
            "country": "India",
            "jobTitle": "VP Mgr-Product Mgmt",
            "presence": null,
            "timeZone": "Asia/Calcutta",
            "department": "TSG",
            "photo": "/_layouts/15/userphoto.aspx?size=L&username=DwivediDi@mcolab.moodys.com",
            "isCoreTeamMember": true
        },
        {
            "id": "9e4a9056-b8ac-4f7a-be25-caecc663dcf5",
            "name": "Alastair Aston",
            "mail": "Alastair.Aston@mcolab.moodys.com",
            "userPrincipalName": "astonal@mcolab.moodys.com",
            "city": "London",
            "country": "United Kingdom",
            "jobTitle": "VP Mgr-Product Mgmt",
            "presence": null,
            "timeZone": "Europe/London",
            "department": "TSG",
            "photo": "/_layouts/15/userphoto.aspx?size=L&username=astonal@mcolab.moodys.com",
            "isCoreTeamMember": true
        },
        {
            "id": "6827819b-6974-4798-9fef-d063256b9b1d",
            "name": "Shobhit Garg",
            "mail": "Shobhit.Garg@mcolab.moodys.com",
            "userPrincipalName": "GargS14@mcolab.moodys.com",
            "city": "Gurugram",
            "country": "India",
            "jobTitle": "Senior Software Engineer",
            "presence": null,
            "timeZone": "Asia/Calcutta",
            "department": "TSG",
            "photo": "/_layouts/15/userphoto.aspx?size=L&username=GargS14@mcolab.moodys.com",
            "isCoreTeamMember": false
        },
        {
            "id": "8bacf7db-aeb6-4197-a02f-24593af0fd49",
            "name": "Deepti Joshi",
            "mail": "Deepti.Joshi@mcolab.moodys.com",
            "userPrincipalName": "joshid@mcolab.moodys.com",
            "city": "Gurugram",
            "country": "India",
            "jobTitle": "Software Engineer",
            "presence": null,
            "timeZone": "Asia/Calcutta",
            "department": "CPG OU",
            "photo": "/_layouts/15/userphoto.aspx?size=L&username=joshid@mcolab.moodys.com",
            "isCoreTeamMember": false
        }
    ]
}