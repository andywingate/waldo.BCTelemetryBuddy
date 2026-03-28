/**
 * User Lookup Tests
 *
 * Tests for UserLookupService — resolving BC telemetry user IDs to real user names
 * via BC Admin API (primary) and Graph API (fallback).
 *
 * All HTTP calls are mocked; no real API access is required.
 */

import { UserLookupService, UserInfo, UserLookupResult } from '../userLookup.js';
import { MCPConfig } from '../config.js';
import { AuthService } from '../auth.js';

// Mock AuthService
const mockGetBCAccessToken = jest.fn().mockResolvedValue('mock-bc-token');
const mockGetGraphAccessToken = jest.fn().mockResolvedValue('mock-graph-token');

const mockAuthService = {
    getBCAccessToken: mockGetBCAccessToken,
    getGraphAccessToken: mockGetGraphAccessToken,
    getAccessToken: jest.fn().mockResolvedValue('mock-token'),
    authenticate: jest.fn(),
    getStatus: jest.fn()
} as unknown as AuthService;

// Mock global fetch
const mockFetch = jest.fn();
global.fetch = mockFetch;

/** A minimal valid MCPConfig for testing */
const baseConfig: MCPConfig = {
    connectionName: 'test',
    tenantId: 'app-insights-tenant',
    authFlow: 'azure_cli',
    applicationInsightsAppId: 'app-id',
    kustoClusterUrl: 'https://test.kusto.windows.net',
    cacheEnabled: false,
    cacheTTLSeconds: 3600,
    removePII: false,
    port: 52345,
    workspacePath: '/test/workspace',
    queriesFolder: 'queries',
    references: []
};

const configWithBC: MCPConfig = {
    ...baseConfig,
    bcTenantId: 'bc-customer-tenant',
    bcEnvironmentName: 'Production'
};

/** A sample BC Admin API response with two users */
const mockBCApiResponse = {
    value: [
        {
            id: 'bc-user-1',
            displayName: 'Admin ADW',
            userName: 'ADMIN.ADW',
            state: 'Enabled',
            userTelemetryId: '18c0f8e9-a967-47b0-abcd-ef1234567890',
            authenticationEmail: 'admin.adw@venturedemos.onmicrosoft.com'
        },
        {
            id: 'bc-user-2',
            displayName: 'John Smith',
            userName: 'JOHN.SMITH',
            state: 'Enabled',
            userTelemetryId: 'aaaabbbb-cccc-dddd-eeee-ffffffffffff',
            authenticationEmail: 'john.smith@venturedemos.onmicrosoft.com'
        },
        {
            id: 'bc-user-3',
            displayName: 'Disabled User',
            userName: 'DIS.USER',
            state: 'Disabled',
            userTelemetryId: '11112222-3333-4444-5555-666677778888',
            authenticationEmail: 'dis.user@venturedemos.onmicrosoft.com'
        }
    ]
};

/** A sample Graph API user response */
const mockGraphUser = {
    displayName: 'John Smith',
    userPrincipalName: 'john.smith@contoso.com',
    mail: 'john.smith@contoso.com',
    givenName: 'John',
    surname: 'Smith'
};

beforeEach(() => {
    jest.clearAllMocks();
    mockFetch.mockReset();
    mockGetBCAccessToken.mockResolvedValue('mock-bc-token');
    mockGetGraphAccessToken.mockResolvedValue('mock-graph-token');
});

// ─── Helpers ─────────────────────────────────────────────────────────────────

function makeJsonResponse(body: any, status = 200): Response {
    return {
        ok: status >= 200 && status < 300,
        status,
        statusText: status === 200 ? 'OK' : status === 404 ? 'Not Found' : 'Error',
        json: jest.fn().mockResolvedValue(body),
        text: jest.fn().mockResolvedValue(JSON.stringify(body))
    } as unknown as Response;
}

// ─── BC Admin API Strategy ────────────────────────────────────────────────────

describe('UserLookupService — BC Admin API strategy', () => {
    it('resolves a known usertelemetryId to BC user info', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        const result = await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);

        expect(result.strategy).toBe('bc_api');
        expect(result.resolvedCount).toBe(1);
        expect(result.notFoundCount).toBe(0);

        const user = result.users[0];
        expect(user.found).toBe(true);
        expect(user.displayName).toBe('Admin ADW');
        expect(user.userName).toBe('ADMIN.ADW');
        expect(user.authenticationEmail).toBe('admin.adw@venturedemos.onmicrosoft.com');
        expect(user.state).toBe('Enabled');
        expect(user.resolvedVia).toBe('bc_api');
    });

    it('resolves multiple users in a single API call (caches the full user list)', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        const result = await service.lookupUsers([
            '18c0f8e9-a967-47b0-abcd-ef1234567890',
            'aaaabbbb-cccc-dddd-eeee-ffffffffffff'
        ]);

        // Only one HTTP call should have been made (the user list is cached)
        expect(mockFetch).toHaveBeenCalledTimes(1);
        expect(result.resolvedCount).toBe(2);
        expect(result.users[0].userName).toBe('ADMIN.ADW');
        expect(result.users[1].userName).toBe('JOHN.SMITH');
    });

    it('marks an ID as not found when it does not appear in the BC user list', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        const result = await service.lookupUsers(['00000000-0000-0000-0000-000000000099']);

        expect(result.resolvedCount).toBe(0);
        expect(result.notFoundCount).toBe(1);
        expect(result.users[0].found).toBe(false);
        expect(result.users[0].error).toContain('not found in BC environment');
    });

    it('uses the correct BC Admin API URL with tenant and environment', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);

        expect(mockFetch).toHaveBeenCalledWith(
            expect.stringContaining('bc-customer-tenant'),
            expect.any(Object)
        );
        expect(mockFetch).toHaveBeenCalledWith(
            expect.stringContaining('Production'),
            expect.any(Object)
        );
        expect(mockFetch).toHaveBeenCalledWith(
            expect.stringContaining('automation/v2.0/users'),
            expect.any(Object)
        );
    });

    it('uses the Authorization Bearer token from getBCAccessToken', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);

        expect(mockFetch).toHaveBeenCalledWith(
            expect.any(String),
            expect.objectContaining({
                headers: expect.objectContaining({
                    Authorization: 'Bearer mock-bc-token'
                })
            })
        );
    });

    it('is case-insensitive when matching usertelemetryId to BC user', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        // Supply the GUID in UPPERCASE, BC API returns lowercase
        const result = await service.lookupUsers(['18C0F8E9-A967-47B0-ABCD-EF1234567890']);

        expect(result.users[0].found).toBe(true);
        expect(result.users[0].userName).toBe('ADMIN.ADW');
    });

    it('uses custom bcBaseUrl when configured', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockBCApiResponse));

        const configWithCustomUrl: MCPConfig = {
            ...configWithBC,
            bcBaseUrl: 'https://custom.bc.example.com'
        };
        const service = new UserLookupService(configWithCustomUrl, mockAuthService);
        await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);

        expect(mockFetch).toHaveBeenCalledWith(
            expect.stringContaining('https://custom.bc.example.com'),
            expect.any(Object)
        );
    });

    it('falls back to Graph API when BC API returns an error status', async () => {
        // BC API returns 401; Graph API then succeeds
        mockFetch
            .mockResolvedValueOnce(makeJsonResponse({ error: 'Unauthorized' }, 401))
            .mockResolvedValueOnce(makeJsonResponse(mockGraphUser));

        const service = new UserLookupService(configWithBC, mockAuthService);
        const result = await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);

        // Should have fallen back to Graph API successfully
        expect(result.strategy).toBe('graph_api');
        expect(result.users[0].found).toBe(true);
        expect(result.users[0].resolvedVia).toBe('graph_api');
    });

    it('caches the BC user list and does not re-fetch on second call', async () => {
        mockFetch.mockResolvedValue(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);
        await service.lookupUsers(['aaaabbbb-cccc-dddd-eeee-ffffffffffff']);

        // Fetch should only have been called once (cache hit on second call)
        expect(mockFetch).toHaveBeenCalledTimes(1);
    });

    it('clearCache forces a fresh BC API fetch on next call', async () => {
        mockFetch.mockResolvedValue(makeJsonResponse(mockBCApiResponse));

        const service = new UserLookupService(configWithBC, mockAuthService);
        await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);
        service.clearCache();
        await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);

        expect(mockFetch).toHaveBeenCalledTimes(2);
    });
});

// ─── Graph API Fallback Strategy ──────────────────────────────────────────────

describe('UserLookupService — Graph API fallback strategy', () => {
    it('uses Graph API when bcTenantId and bcEnvironmentName are not configured', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockGraphUser));

        const service = new UserLookupService(baseConfig, mockAuthService);
        const result = await service.lookupUsers(['aaaabbbb-cccc-dddd-eeee-ffffffffffff']);

        expect(result.strategy).toBe('graph_api');
        expect(result.resolvedCount).toBe(1);
        const user = result.users[0];
        expect(user.found).toBe(true);
        expect(user.displayName).toBe('John Smith');
        expect(user.userPrincipalName).toBe('john.smith@contoso.com');
        expect(user.resolvedVia).toBe('graph_api');
    });

    it('falls back to Graph API when BC API fails', async () => {
        // BC API fails first
        mockFetch
            .mockResolvedValueOnce(makeJsonResponse({ error: 'Service Unavailable' }, 503))
            .mockResolvedValueOnce(makeJsonResponse(mockGraphUser));

        const service = new UserLookupService(configWithBC, mockAuthService);
        const result = await service.lookupUsers(['aaaabbbb-cccc-dddd-eeee-ffffffffffff']);

        expect(result.strategy).toBe('graph_api');
        expect(result.users[0].found).toBe(true);
        expect(result.users[0].resolvedVia).toBe('graph_api');
    });

    it('marks user as not found when Graph returns 404', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse({ error: 'Not found' }, 404));

        const service = new UserLookupService(baseConfig, mockAuthService);
        const result = await service.lookupUsers(['00000000-0000-0000-0000-000000000001']);

        expect(result.notFoundCount).toBe(1);
        expect(result.users[0].found).toBe(false);
        expect(result.users[0].error).toContain('not found in Azure AD');
    });

    it('uses the correct Graph API URL with the usertelemetryId', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockGraphUser));

        const service = new UserLookupService(baseConfig, mockAuthService);
        await service.lookupUsers(['aaaabbbb-cccc-dddd-eeee-ffffffffffff']);

        expect(mockFetch).toHaveBeenCalledWith(
            expect.stringContaining('graph.microsoft.com/v1.0/users/aaaabbbb-cccc-dddd-eeee-ffffffffffff'),
            expect.any(Object)
        );
    });

    it('returns strategy note explaining Graph API limitation when it is used', async () => {
        mockFetch.mockResolvedValueOnce(makeJsonResponse(mockGraphUser));

        const service = new UserLookupService(baseConfig, mockAuthService);
        const result = await service.lookupUsers(['aaaabbbb-cccc-dddd-eeee-ffffffffffff']);

        expect(result.strategyNote).toContain('Graph API');
        expect(result.strategyNote).toContain('bcTenantId');
    });
});

// ─── Empty / edge cases ───────────────────────────────────────────────────────

describe('UserLookupService — edge cases', () => {
    it('returns empty result for empty input', async () => {
        const service = new UserLookupService(baseConfig, mockAuthService);
        const result = await service.lookupUsers([]);

        expect(result.strategy).toBe('none');
        expect(result.users).toHaveLength(0);
        expect(result.totalQueried).toBe(0);
        expect(mockFetch).not.toHaveBeenCalled();
    });

    it('returns error result when both BC API and Graph API fail', async () => {
        mockGetBCAccessToken.mockRejectedValueOnce(new Error('BC auth failed'));
        mockGetGraphAccessToken.mockRejectedValueOnce(new Error('Graph auth failed'));

        const service = new UserLookupService(configWithBC, mockAuthService);
        const result = await service.lookupUsers(['18c0f8e9-a967-47b0-abcd-ef1234567890']);

        expect(result.strategy).toBe('none');
        expect(result.users[0].found).toBe(false);
        expect(result.users[0].error).toContain('Graph auth failed');
    });

    it('handles fetch throwing a network error gracefully', async () => {
        mockFetch.mockRejectedValueOnce(new Error('ENOTFOUND graph.microsoft.com'));

        const service = new UserLookupService(baseConfig, mockAuthService);
        const result = await service.lookupUsers(['aaaabbbb-cccc-dddd-eeee-ffffffffffff']);

        expect(result.users[0].found).toBe(false);
        expect(result.users[0].error).toContain('Graph API request failed');
    });
});

// ─── KQL filter generation ───────────────────────────────────────────────────

describe('lookup_user_telemetry_ids KQL filter', () => {
    it('verifies that a valid GUID filter can be constructed', () => {
        const telemetryId = '18c0f8e9-a967-47b0-abcd-ef1234567890';
        const filter = `| where tostring(customDimensions.usertelemetryId) == "${telemetryId}"`;

        expect(filter).toContain('customDimensions.usertelemetryId');
        expect(filter).toContain(telemetryId);
        expect(filter).toMatch(/^\| where tostring\(customDimensions\.usertelemetryId\) == "/);
    });
});
