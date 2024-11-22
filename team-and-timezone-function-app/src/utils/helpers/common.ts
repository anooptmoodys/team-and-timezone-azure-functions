import { findIana } from "windows-iana";

// function to extract upn from token
export function extractUpnFromToken(token: string): string | null {
    const tokenPayload = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString());
    return tokenPayload.upn || null;
}

// function to get the IANA time zone from the Windows time zone
export function getIanaFromWindows(windowsTimeZone: string): string {
    const ianaZones = findIana(windowsTimeZone);
    if(!ianaZones || ianaZones.length === 0) {
        return "Europe/London";
    }
    return ianaZones[0];
}