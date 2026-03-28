/**
 * User Lookup Service
 *
 * Resolves Business Central telemetry user IDs (usertelemetryId) to real user names
 * using the Microsoft Graph API. This is useful for identifying which user triggered
 * a specific telemetry event — particularly permission errors.
 *
 * Prerequisites:
 * - The App Insights instance and the BC environment must be in the same AAD tenant.
 * - The authenticated identity must have permission to read users from Microsoft Graph
 *   (User.Read.All or Directory.Read.All for service principals; User.Read for delegated).
 */

import { AuthService } from './auth.js';
import { MCPConfig } from './config.js';

/**
 * Resolved user information from Microsoft Graph API
 */
export interface UserInfo {
    /** The original usertelemetryId from BC telemetry (AAD Object ID) */
    usertelemetryId: string;
    /** Display name (e.g. "John Smith") */
    displayName?: string;
    /** User Principal Name / login (e.g. "john.smith@contoso.com") */
    userPrincipalName?: string;
    /** Email address (may differ from UPN in some tenants) */
    mail?: string;
    /** Given name */
    givenName?: string;
    /** Surname */
    surname?: string;
    /** Whether the user was successfully found in Azure AD */
    found: boolean;
    /** Error message if the lookup failed */
    error?: string;
}

/**
 * Service that maps BC telemetry user IDs to real user information via Microsoft Graph API.
 *
 * usertelemetryId in modern BC (2021+) is the AAD Object ID of the user, which allows
 * a direct lookup via https://graph.microsoft.com/v1.0/users/{objectId}.
 */
export class UserLookupService {
    private config: MCPConfig;
    private authService: AuthService;
    /** In-memory cache to avoid repeated Graph API calls within the same MCP session */
    private readonly userCache = new Map<string, UserInfo>();

    constructor(config: MCPConfig, authService: AuthService) {
        this.config = config;
        this.authService = authService;
    }

    /**
     * Look up a single user by their BC telemetry ID (AAD Object ID).
     * Results are cached per session to minimise Graph API calls.
     */
    async lookupUser(usertelemetryId: string): Promise<UserInfo> {
        if (this.userCache.has(usertelemetryId)) {
            return this.userCache.get(usertelemetryId)!;
        }

        let graphToken: string;
        try {
            graphToken = await this.authService.getGraphAccessToken();
        } catch (error: any) {
            const info: UserInfo = {
                usertelemetryId,
                found: false,
                error: `Cannot obtain Graph API token: ${error.message}`
            };
            return info;
        }

        try {
            const url =
                `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(usertelemetryId)}` +
                '?$select=displayName,userPrincipalName,mail,givenName,surname';

            const response = await fetch(url, {
                headers: {
                    Authorization: `Bearer ${graphToken}`,
                    'Content-Type': 'application/json'
                }
            });

            if (!response.ok) {
                const errorText = await response.text().catch(() => '');
                const info: UserInfo = {
                    usertelemetryId,
                    found: false,
                    error: response.status === 404
                        ? 'User not found in Azure AD directory (may be from a different tenant or deleted)'
                        : `Graph API error ${response.status}: ${response.statusText}${errorText ? ` — ${errorText}` : ''}`
                };
                this.userCache.set(usertelemetryId, info);
                return info;
            }

            const user = await response.json() as Record<string, string | null>;
            const info: UserInfo = {
                usertelemetryId,
                displayName: user.displayName ?? undefined,
                userPrincipalName: user.userPrincipalName ?? undefined,
                mail: user.mail ?? undefined,
                givenName: user.givenName ?? undefined,
                surname: user.surname ?? undefined,
                found: true
            };

            this.userCache.set(usertelemetryId, info);
            return info;
        } catch (error: any) {
            const info: UserInfo = {
                usertelemetryId,
                found: false,
                error: `Graph API request failed: ${error.message}`
            };
            return info;
        }
    }

    /**
     * Look up multiple users in parallel.
     * Each lookup is independent; failures are captured per-entry.
     */
    async lookupUsers(usertelemetryIds: string[]): Promise<UserInfo[]> {
        return Promise.all(usertelemetryIds.map(id => this.lookupUser(id)));
    }

    /**
     * Clear the in-session user cache (useful for testing or profile switches).
     */
    clearCache(): void {
        this.userCache.clear();
    }
}
