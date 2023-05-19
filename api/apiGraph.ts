// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// import fetch from 'node-fetch';
import { OnBehalfOfUserCredential, createMicrosoftGraphClientWithCredential } from '@microsoft/teamsfx';
import AuthConfig from '../config/authConfig';
import { Client, ResponseType } from '@microsoft/microsoft-graph-client';

// empty image to default to
const emptyPic = "data:image/jpeg;base64, iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAA51JREFUWAnFV0tIlFEU/u4/5oDOjM1YECmJkloKPQzNpYGhlAXVqgwqpEUuI1xUtNDaSquCnJJKdz1AaF1LH2X2EnwsFEYNAh/zzwwUOLdz7jz4Z+b3/8cX/sPM3P+cc8/33XPvPfdcgSwfKWVuOBw+JeXqeQFRJSX2A5K+/Ih5ITAvIceFcAzk5+d/FEL8i+msf4W1GgiFQvsI9AEkWsnWY2cf1wch0E9kOl0u12+rPmsSoBE7Q6HgPSFxWwL5Vk7W0pHzsBTodrk8jygif83sTAmoUUej7ynE9Wad1i8Tg0LTLphFI4MAgR+BjH6gCBSvH2jtHhSBAIR2hkj8MFqlEOCRE/jIVoMnAOMkao2R0BJKnnNJYd8ucMZh33EMZwI3SYAX3NbNecK92b+sj2HFdGoKVOijq9MbXe1mMFYyAg1DcxzkqchhQ7XP17HVIpEw+l69xsiXz7TdBerqanGl9Sry8vKscJM6NVDOLUC7oHnJDenBP/SSVZIhe9y/fxdlZQdx6eJFyk8S7968xczMDDofPgQttCSQTSPocnv2apxeswVnh19HR+EQGtra2rDb64XX60PbzZuI0ufb2JgNZoraw9ga5/YUsc1LIBBAeWVlhlV5eQXm5uYy5FYCxtb4YLEyStcVFxdjamIiXYypyUkUFRVlyK0EjE0R4FMt++d4TQ0FW8Lv92N5aQmLi4vw9zyDQ9Nw9Nix7B2RJWMLPbiiU9u1np6RSAT9fX0YHhlWu+BkXR0ut7ZmvQsMWKENETA42GxTp0wo5jfrZeP9xUKOqmQkKqyc3Lh+jQsTK5MMXYHHA/+L3gy5UcDYOVxGkbDBqEhv+3w+3OnoQHVVdbrK9D0wF8Cjri5TnVHI2BrXcEahWbul5Rye9/gxSVvN7pmdncXj7m40NTfbmVLWdAxknYqHhofwsrcXXp8Xhw4dRsmBErgpzHS8Ynl5BQsL87QrhpC7y0ngTWhsPG1HQKVilbh1feUJbe1bdj2iBDY9NY1f4z/BGVEPBuFwOFQ6LtxTiJqaEygtLbVzE9MLPHW7C9oVgZ08jlVBwucyV6/ZUd+8VaxSjpXrKgLsko5ZZ0jXP21/VSQGXW53Q6JMT5ZkLODSWRWOmx+kqQf2HcdI3hGSBLiHqlaFdnY7SCifsbI85aaUQiBO4jvV77WUogdNh7EhIfkinzTAlDsBu0qugXS/ak3s1NXMSIa36I5cTo0kuM3F63Zcz/8DmbGCUuUyJvcAAAAASUVORK5CYII=";


/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
class ApiGraph {
    private graphClient: Client;

    constructor(accessToken: string) {
        if (!accessToken || !accessToken.trim()) {
            throw new Error('[ApiGraph]: Missing SSO token.');
        }

        // Call Microsoft Graph on behalf of user
        const oboCredential = new OnBehalfOfUserCredential(accessToken, AuthConfig.oboAuthConfig);
        this.graphClient = createMicrosoftGraphClientWithCredential(oboCredential, [
            "User.Read",
        ]);
    }

    
    /**
     * Collects information about the user in the bot.
     */
    public getPersonAsync = async (userId?: string | undefined): Promise<any> => {
        const apiUrl = (userId ? `/users/${userId}` : '/me');

        try {
            const user = await this.graphClient.api(apiUrl).get();
            return user;
        } catch (error) {
            if(error instanceof Error) {
                // Something happened in setting up the request that triggered an Error
                console.log('Error', error.message);
            }
            throw error;
        }
    };    

    // Gets the user's photo
    public getUserPhotoAsync = async (
        userId?: string | undefined,
    ): Promise<string> => {
        const apiUri = (userId ? `/users/${userId}/photo/$value` : '/me/photo/$value');

        let photoBinary: ArrayBuffer;
        try {
          photoBinary = await this.graphClient
            .api(apiUri)
            .responseType(ResponseType.ARRAYBUFFER)
            .get();
        } catch (error) {
            if(error instanceof Error) {
                // Something happened in setting up the request that triggered an Error
                console.log('Error', error.message);
            }
            return emptyPic;
        }
  
        const buffer = Buffer.from(photoBinary);
        const pic = "data:image/png;base64," + buffer.toString("base64");
  
        return pic;
    };

}

export default ApiGraph;