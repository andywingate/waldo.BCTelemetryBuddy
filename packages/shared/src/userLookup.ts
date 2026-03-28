/**
 * User Lookup Service
 *
 * Resolves Business Central telemetry user IDs (`customDimensions.usertelemetryId`)
 * to real, human-readable user names.
 *
 * ## Why this is needed
 * BC telemetry intentionally records a GUID (`usertelemetryId`) rather than a name to
 * avoid embedding PII in telemetry by default. In practice, admins and partners often
 * need to know **which user** triggered a specific event (especially permission errors).
 * This service automates that reverse-lookup.
 *
 * ## Two lookup strategies (tried in order)
 *
 * ### Strategy A — BC Admin API (primary, most accurate)
 * Requires: `bcTenantId` + `bcEnvironmentName` in config.
 *
 * Queries BC's own Automation API:
 *   `GET /v2.0/{bcTenantId}/{bcEnvironmentName}/api/microsoft/automation/v2.0/users`
 *
 * This returns every user in the environment with their `userTelemetryId`, `userName`,
 * `displayName`, and `authenticationEmail` — the authoritative mapping straight from BC.
 *
 * This is the right strategy for the common cross-tenant case where App Insights lives
 * in an ISV/partner tenant and BC runs in the customer's tenant.
 *
 * ### Strategy B — Microsoft Graph API (fallback)
 * In BC SaaS (cloud), `usertelemetryId` equals the user's AAD Object ID.
 * When BC and App Insights are in the **same** AAD tenant, a Graph lookup works:
 *   `GET https://graph.microsoft.com/v1.0/users/{usertelemetryId}`
 *
 * This does NOT require `bcTenantId`/`bcEnvironmentName` but only works for same-tenant
 * deployments, and relies on the assumption that usertelemetryId == AAD Object ID.
 *
 * ## PII / consent note
 * Resolving user names from telemetry exposes PII. This service assumes the end-user
 * organization has consented to this mapping — consent is managed outside the MCP server.
 */

import { AuthService } from './auth.js';
import { MCPConfig } from './config.js';

const BC_DEFAULT_BASE_URL = 'https://api.businesscentral.dynamics.com';

/**
 * A BC Automation API user record
 */
interface BCApiUser {
    id: string;
    displayName: string;
    userName: string;
    state: string;
    userTelemetryId: string;
    authenticationEmail?: string;
}

/**
 * Resolved user information — result of looking up a `usertelemetryId`
 */
export interface UserInfo {
    /** The `usertelemetryId` value from BC telemetry (`customDimensions.usertelemetryId`) */
    usertelemetryId: string;
    /** Display name (e.g. "John Smith") */
    displayName?: string;
    /** BC user name / login (e.g. "ADMIN.ADW") */
    userName?: string;
    /** Authentication email — Microsoft 365 email for SaaS users */
    authenticationEmail?: string;
    /** User Principal Name from Graph API (may equal authenticationEmail) */
    userPrincipalName?: string;
    /** User account state in BC ("Enabled" / "Disabled") */
    state?: string;
    /** Which lookup strategy successfully resolved this user */
    resolvedVia?: 'bc_api' | 'graph_api';
    /** Whether the user was successfully resolved */
    found: boolean;
    /** Error or reason when `found` is false */
    error?: string;
}

/**
 * Lookup result summary returned to the MCP tool caller
 */
export interface UserLookupResult {
    strategy: 'bc_api' | 'graph_api' | 'none';
    users: UserInfo[];
    totalQueried: number;
    resolvedCount: number;
    notFoundCount: number;
    strategyNote: string;
}

export class UserLookupService {
    private config: MCPConfig;
    private authService: AuthService;

    /** Session-level cache: usertelemetryId → resolved UserInfo */
    private readonly userCache = new Map<string, UserInfo>();

    /** Session-level cache of the full BC user list (populated on first BC API call) */
    private bcUserMap: Map<string, BCApiUser> | null = null;

    constructor(config: MCPConfig, authService: AuthService) {
        this.config = config;
        this.authService = authService;
    }

    /**
     * Resolve a list of `usertelemetryId` values to real user names.
     *
     * Tries the BC Admin API first (authoritative), falls back to Graph API.
     * Results are cached within the MCP session to minimise API calls.
     */
    async lookupUsers(usertelemetryIds: string[]): Promise<UserLookupResult> {
        if (usertelemetryIds.length === 0) {
            return this.emptyResult('none', 'No usertelemetryId values provided.');
        }

        // Strategy A: BC Admin API
        const hasBCConfig = !!(this.config.bcTenantId && this.config.bcEnvironmentName);
        if (hasBCConfig) {
            try {
                return await this.lookupViaBCApi(usertelemetryIds);
            } catch (bcError: any) {
                // Log the error but fall through to Graph API
                console.error(`[UserLookup] BC API lookup failed: ${bcError.message}. Trying Graph API fallback.`);
            }
        }

        // Strategy B: Graph API fallback
        try {
            return await this.lookupViaGraphApi(usertelemetryIds);
        } catch (graphError: any) {
            // Both strategies failed — return informative error result
            const users: UserInfo[] = usertelemetryIds.map(id => ({
                usertelemetryId: id,
                found: false,
                error: graphError.message
            }));
            return {
                strategy: 'none',
                users,
                totalQueried: usertelemetryIds.length,
                resolvedCount: 0,
                notFoundCount: users.length,
                strategyNote: this.buildStrategyNote(hasBCConfig, false, false, graphError.message)
            };
        }
    }

    // ─── Strategy A: BC Admin API ─────────────────────────────────────────────

    private async lookupViaBCApi(usertelemetryIds: string[]): Promise<UserLookupResult> {
        const bcMap = await this.getBCUserMap();

        const users: UserInfo[] = usertelemetryIds.map(id => {
            if (this.userCache.has(id)) return this.userCache.get(id)!;

            const bcUser = bcMap.get(id.toLowerCase());
            if (bcUser) {
                const info: UserInfo = {
                    usertelemetryId: id,
                    displayName: bcUser.displayName || undefined,
                    userName: bcUser.userName || undefined,
                    authenticationEmail: bcUser.authenticationEmail || undefined,
                    state: bcUser.state || undefined,
                    found: true,
                    resolvedVia: 'bc_api'
                };
                this.userCache.set(id, info);
                return info;
            }

            const notFound: UserInfo = {
                usertelemetryId: id,
                found: false,
                error: 'usertelemetryId not found in BC environment user list'
            };
            this.userCache.set(id, notFound);
            return notFound;
        });

        const resolved = users.filter(u => u.found).length;
        return {
            strategy: 'bc_api',
            users,
            totalQueried: usertelemetryIds.length,
            resolvedCount: resolved,
            notFoundCount: users.length - resolved,
            strategyNote: `Resolved via BC Admin API (${this.config.bcEnvironmentName} / ${this.config.bcTenantId}). This is the authoritative source — user names come directly from BC's User table.`
        };
    }

    /**
     * Fetch all users from BC Admin API and build a Map keyed by lowercase usertelemetryId.
     * Result is cached for the lifetime of the service instance (one MCP session).
     */
    private async getBCUserMap(): Promise<Map<string, BCApiUser>> {
        if (this.bcUserMap) return this.bcUserMap;

        const token = await this.authService.getBCAccessToken();
        const baseUrl = this.config.bcBaseUrl || BC_DEFAULT_BASE_URL;
        const tenantId = this.config.bcTenantId!;
        const environmentName = encodeURIComponent(this.config.bcEnvironmentName!);

        // Automation API: list all users with their telemetryId
        const url = `${baseUrl}/v2.0/${tenantId}/${environmentName}/api/microsoft/automation/v2.0/users`;

        const response = await fetch(url, {
            headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            const body = await response.text().catch(() => '');
            throw new Error(
                `BC Admin API error ${response.status}: ${response.statusText}` +
                (body ? ` — ${body.slice(0, 200)}` : '')
            );
        }

        const data = await response.json() as { value?: BCApiUser[] };
        const bcUsers: BCApiUser[] = data.value || [];

        const map = new Map<string, BCApiUser>();
        for (const user of bcUsers) {
            if (user.userTelemetryId) {
                map.set(user.userTelemetryId.toLowerCase(), user);
            }
        }

        this.bcUserMap = map;
        return map;
    }

    // ─── Strategy B: Graph API ─────────────────────────────────────────────────

    private async lookupViaGraphApi(usertelemetryIds: string[]): Promise<UserLookupResult> {
        const token = await this.authService.getGraphAccessToken();

        const users = await Promise.all(
            usertelemetryIds.map(id => this.lookupSingleViaGraph(id, token))
        );

        const resolved = users.filter(u => u.found).length;
        return {
            strategy: 'graph_api',
            users,
            totalQueried: usertelemetryIds.length,
            resolvedCount: resolved,
            notFoundCount: users.length - resolved,
            strategyNote: 'Resolved via Microsoft Graph API. This works when BC and App Insights share the same AAD tenant and usertelemetryId equals the AAD Object ID (true for BC SaaS). For cross-tenant deployments, configure bcTenantId and bcEnvironmentName to use the authoritative BC Admin API instead.'
        };
    }

    private async lookupSingleViaGraph(usertelemetryId: string, token: string): Promise<UserInfo> {
        if (this.userCache.has(usertelemetryId)) {
            return this.userCache.get(usertelemetryId)!;
        }

        try {
            const url =
                `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(usertelemetryId)}` +
                '?$select=displayName,userPrincipalName,mail,givenName,surname';

            const response = await fetch(url, {
                headers: {
                    Authorization: `Bearer ${token}`,
                    'Content-Type': 'application/json'
                }
            });

            if (!response.ok) {
                const info: UserInfo = {
                    usertelemetryId,
                    found: false,
                    error: response.status === 404
                        ? 'User not found in Azure AD (may be from a different tenant, or usertelemetryId may not equal AAD Object ID for this BC deployment)'
                        : `Graph API error ${response.status}: ${response.statusText}`
                };
                this.userCache.set(usertelemetryId, info);
                return info;
            }

            const user = await response.json() as Record<string, string | null>;
            const info: UserInfo = {
                usertelemetryId,
                displayName: user.displayName ?? undefined,
                userPrincipalName: user.userPrincipalName ?? undefined,
                authenticationEmail: user.mail ?? undefined,
                found: true,
                resolvedVia: 'graph_api'
            };
            this.userCache.set(usertelemetryId, info);
            return info;
        } catch (error: any) {
            const info: UserInfo = {
                usertelemetryId,
                found: false,
                error: `Graph API request failed: ${error.message}`
            };
            this.userCache.set(usertelemetryId, info);
            return info;
        }
    }

    // ─── Helpers ───────────────────────────────────────────────────────────────

    private emptyResult(strategy: 'none', note: string): UserLookupResult {
        return {
            strategy,
            users: [],
            totalQueried: 0,
            resolvedCount: 0,
            notFoundCount: 0,
            strategyNote: note
        };
    }

    private buildStrategyNote(
        hasBCConfig: boolean,
        bcSucceeded: boolean,
        graphSucceeded: boolean,
        errorMsg?: string
    ): string {
        const parts: string[] = [];
        if (!hasBCConfig) {
            parts.push('BC Admin API not configured (bcTenantId and bcEnvironmentName are required for the authoritative lookup).');
        }
        if (!graphSucceeded && errorMsg) {
            parts.push(`Graph API fallback also failed: ${errorMsg}`);
        }
        parts.push(
            'To enable user name lookup: ' +
            (hasBCConfig
                ? 'check that the configured BC credentials have API access.'
                : 'set BCTB_BC_TENANT_ID and BCTB_BC_ENVIRONMENT_NAME in your BCTB config.')
        );
        return parts.join(' ');
    }

    /**
     * Clear all cached user data (useful when switching profiles or in tests)
     */
    clearCache(): void {
        this.userCache.clear();
        this.bcUserMap = null;
    }
}
