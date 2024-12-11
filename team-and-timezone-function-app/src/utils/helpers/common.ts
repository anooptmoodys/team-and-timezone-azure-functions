import { findIana } from "windows-iana";

// function to get the IANA time zone from the Windows time zone
export function getIanaFromWindows(windowsTimeZone: string): string {
    const ianaZones = findIana(windowsTimeZone);
    if(!ianaZones || ianaZones.length === 0) {
        return "Europe/London";
    }
    return ianaZones[0];
}

// function to return error response
/* 
Sample error response:
{
    status: 401,
    jsonBody: {
        status: 401,
        error: {
            exists: true,
            code: "Unauthorized",
            message: "Unauthorized"
        },
        data: null
    }
};
 */
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