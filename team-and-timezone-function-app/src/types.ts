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
    department?: string;
    isCoreTeamMember?: boolean; //needed only for initial load to separate core team members from people working with
}

export type TokenValidationResult = {
    valid: boolean;
    errorMessage: string
}