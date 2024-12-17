import { User } from "@microsoft/microsoft-graph-types";
import { TeamMember } from "../../types";
import { getIanaFromWindows } from "./common";

export const mapTeamMemberDetails = (teamMember: User, presenceData: any, timezoneData: any): TeamMember => {
    return {
        id: teamMember.id,
        name: teamMember.givenName && teamMember.surname ? `${teamMember.givenName} ${teamMember.surname}` : teamMember.displayName || null,
        mail: teamMember.mail || null,
        userPrincipalName: teamMember.userPrincipalName || null,
        city: teamMember.city || null,
        country: teamMember.country || null,
        jobTitle: teamMember.jobTitle || null,
        presence: presenceData[teamMember.id] || null,
        timeZone: getIanaFromWindows(timezoneData[teamMember.id] ?? "GMT Standard Time"),
        department: teamMember.department || null,
        photo: `/_layouts/15/userphoto.aspx?size=L&username=${teamMember.userPrincipalName}`
    };
};