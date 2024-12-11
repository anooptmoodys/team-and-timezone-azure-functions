import { TokenValidator, getEntraJwksUri, ValidateTokenOptions } from "jwt-validate";
import { TokenValidationResult } from "../../types";

// function to validate token
export async function validateToken(token: string): Promise<TokenValidationResult> {

    const jwksUri = await getEntraJwksUri();

    const tokenValidator = new TokenValidator({
        jwksUri
    });

    const audience = process.env.CLIENT_ID;
    const issuer = process.env.ENTRA_ISSUER;

    // if audience and issuer are not provided, return false
    if (!audience || !issuer) {
        console.error("Audience and issuer are required to validate token");
        return { valid: false, errorMessage: "Audience and issuer are required to validate token" };
    }

    const options: ValidateTokenOptions = {
        audience,
        issuer
    }
    
    try {
        const validToken = await tokenValidator.validateToken(token, options);
        console.log("Token is valid: ", validToken);
        return { valid: true, errorMessage: "" };
    } catch (error) {
        console.error("Error validating token: ", error);
        return { valid: false, errorMessage: error.message };
    }
}

// function to extract upn from token
export function extractUpnFromToken(token: string): string | null {
    const tokenPayload = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString());
    return tokenPayload.upn || null;
}