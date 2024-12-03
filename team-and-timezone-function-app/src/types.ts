export type TeamMember = {
    id: string;
    name: string;
    mail: string;
    userPrincipalName: string;
    city: string;
    country: string;
    jobTitle: string;
    presence: string;
    timeZone: string;
    photo: string;
    isOtherTeamMember?: boolean;
}