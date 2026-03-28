/**
 * User Lookup Tool Handler Tests
 *
 * Tests for the `lookup_user_telemetry_ids` MCP tool — validates the dispatch,
 * telemetry auto-discovery, GUID validation, and result structure.
 *
 * All external calls (Kusto, BC API, Graph API) are mocked.
 */

// ─── Mocks ────────────────────────────────────────────────────────────────────

// Mock UserLookupResult shape returned by UserLookupService.lookupUsers()
const mockLookupUsers = jest.fn();

jest.mock('@bctb/shared', () => ({
    AuthService: jest.fn().mockImplementation(() => ({
        getStatus: jest.fn().mockReturnValue({ authenticated: false }),
        getAccessToken: jest.fn().mockResolvedValue('mock-token'),
        getBCAccessToken: jest.fn().mockResolvedValue('mock-bc-token'),
        getGraphAccessToken: jest.fn().mockResolvedValue('mock-graph-token')
    })),
    KustoService: jest.fn().mockImplementation(() => ({
        executeQuery: jest.fn().mockResolvedValue({ tables: [] }),
        validateQuery: jest.fn().mockReturnValue([])
    })),
    CacheService: jest.fn().mockImplementation(() => ({
        get: jest.fn().mockReturnValue(null),
        set: jest.fn(),
        clear: jest.fn(),
        cleanupExpired: jest.fn(),
        getStats: jest.fn().mockReturnValue({ totalEntries: 0 })
    })),
    QueriesService: jest.fn().mockImplementation(() => ({
        getAllQueries: jest.fn().mockReturnValue([]),
        searchQueries: jest.fn().mockReturnValue([]),
        saveQuery: jest.fn().mockReturnValue('/tmp/test.kql'),
        getCategories: jest.fn().mockReturnValue([])
    })),
    ReferencesService: jest.fn().mockImplementation(() => ({
        getAllExternalQueries: jest.fn().mockResolvedValue([])
    })),
    UserLookupService: jest.fn().mockImplementation(() => ({
        lookupUsers: mockLookupUsers,
        clearCache: jest.fn()
    })),
    sanitizeObject: jest.fn((x: any) => x),
    lookupEventCategory: jest.fn(),
    NoOpUsageTelemetry: jest.fn().mockImplementation(() => ({
        trackEvent: jest.fn(),
        trackException: jest.fn(),
        flush: jest.fn()
    })),
    RateLimitedUsageTelemetry: jest.fn().mockImplementation(() => ({
        trackEvent: jest.fn(),
        trackException: jest.fn(),
        flush: jest.fn()
    })),
    IUsageTelemetry: jest.fn(),
    TELEMETRY_CONNECTION_STRING: '',
    TELEMETRY_EVENTS: {
        MCP: { ERROR: 'MCP.Error', SERVER_STARTED: 'MCP.ServerStarted' },
        MCP_TOOLS: { QUERY_TELEMETRY: 'Mcp.ToolCompleted' }
    },
    createCommonProperties: jest.fn().mockReturnValue({}),
    cleanTelemetryProperties: jest.fn((x: any) => x),
    hashValue: jest.fn().mockReturnValue('hash')
}));

jest.mock('../mcpTelemetry.js', () => ({
    createMCPUsageTelemetry: jest.fn().mockReturnValue(null),
    getMCPInstallationId: jest.fn().mockReturnValue('test-install-id')
}));

jest.mock('../version.js', () => ({ VERSION: '0.0.0-test' }));

// ─── Subject ─────────────────────────────────────────────────────────────────

import { ToolHandlers, initializeServices } from '../tools/toolHandlers.js';
import { MCPConfig } from '../config.js';

const baseConfig: MCPConfig = {
    connectionName: 'test',
    tenantId: 'app-insights-tenant',
    authFlow: 'azure_cli',
    applicationInsightsAppId: 'test-app-id',
    kustoClusterUrl: 'https://test.kusto.windows.net',
    cacheEnabled: false,
    cacheTTLSeconds: 3600,
    removePII: false,
    port: 52345,
    workspacePath: '/tmp/test-workspace',
    queriesFolder: 'queries',
    references: []
};

const configWithBC: MCPConfig = {
    ...baseConfig,
    bcTenantId: 'bc-customer-tenant',
    bcEnvironmentName: 'Production'
};

/** Successful lookup result from UserLookupService */
const successfulLookupResult = {
    strategy: 'bc_api',
    strategyNote: 'Resolved via BC Admin API.',
    totalQueried: 1,
    resolvedCount: 1,
    notFoundCount: 0,
    users: [
        {
            usertelemetryId: '18c0f8e9-a967-47b0-abcd-ef1234567890',
            displayName: 'Admin ADW',
            userName: 'ADMIN.ADW',
            authenticationEmail: 'admin.adw@venturedemos.onmicrosoft.com',
            state: 'Enabled',
            found: true,
            resolvedVia: 'bc_api'
        }
    ]
};

function createHandlers(config = configWithBC): ToolHandlers {
    const services = initializeServices(config, true);
    return new ToolHandlers(config, services, true, []);
}

beforeEach(() => {
    jest.clearAllMocks();
    mockLookupUsers.mockResolvedValue(successfulLookupResult);
});

// ─── Explicit usertelemetryId input ──────────────────────────────────────────

describe('lookup_user_telemetry_ids tool — explicit IDs', () => {
    it('resolves provided GUIDs and returns enriched user info', async () => {
        const handlers = createHandlers();

        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['18c0f8e9-a967-47b0-abcd-ef1234567890']
        });

        expect(result.resolvedCount).toBe(1);
        expect(result.users[0].displayName).toBe('Admin ADW');
        expect(result.users[0].userName).toBe('ADMIN.ADW');
        expect(result.users[0].kqlFilter).toContain('customDimensions.usertelemetryId');
        expect(result.users[0].kqlFilter).toContain('18c0f8e9-a967-47b0-abcd-ef1234567890');
    });

    it('includes strategy and strategyNote in the result', async () => {
        const handlers = createHandlers();

        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['18c0f8e9-a967-47b0-abcd-ef1234567890']
        });

        expect(result.strategy).toBe('bc_api');
        expect(result.strategyNote).toBeTruthy();
    });

    it('rejects non-GUID usertelemetryIds with a clear error', async () => {
        const handlers = createHandlers();

        await expect(handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['not-a-guid', 'also-not-a-guid']
        })).rejects.toThrow('Invalid usertelemetryId format');
    });

    it('accepts valid GUIDs in any case', async () => {
        const handlers = createHandlers();

        // Should not throw — uppercase is valid GUID format
        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['18C0F8E9-A967-47B0-ABCD-EF1234567890']
        });

        expect(mockLookupUsers).toHaveBeenCalledWith(['18C0F8E9-A967-47B0-ABCD-EF1234567890']);
    });

    it('respects maxUsers cap (100)', async () => {
        const handlers = createHandlers();
        // Supply 150 GUIDs; only 100 should be passed to lookupUsers
        const ids = Array.from({ length: 150 }, (_, i) =>
            `${i.toString(16).padStart(8, '0')}-0000-0000-0000-000000000000`
        );

        await handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ids
        });

        expect(mockLookupUsers).toHaveBeenCalledWith(expect.any(Array));
        const calledWith = mockLookupUsers.mock.calls[0][0] as string[];
        expect(calledWith.length).toBeLessThanOrEqual(100);
    });
});

// ─── Auto-discovery from telemetry ───────────────────────────────────────────

describe('lookup_user_telemetry_ids tool — auto-discovery', () => {
    it('queries telemetry for permission errors when no IDs are provided', async () => {
        const handlers = createHandlers();

        // Mock executeQuery to return a row with a usertelemetryId
        const kusto = (handlers.services as any).kusto;
        kusto.executeQuery = jest.fn().mockResolvedValue({
            type: 'table',
            rows: [['18c0f8e9-a967-47b0-abcd-ef1234567890', 5, '2024-01-01']],
            columns: ['usertelemetryId', 'eventCount', 'lastSeen'],
            summary: '1 row',
            cached: false
        });

        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {});

        expect(kusto.executeQuery).toHaveBeenCalled();
        const kql: string = kusto.executeQuery.mock.calls[0][0];
        expect(kql).toContain('usertelemetryId');
        expect(kql).toContain('AL0000E24'); // default permission error event
    });

    it('uses the provided eventId in the telemetry KQL', async () => {
        const handlers = createHandlers();
        const kusto = (handlers.services as any).kusto;
        kusto.executeQuery = jest.fn().mockResolvedValue({
            type: 'table',
            rows: [['18c0f8e9-a967-47b0-abcd-ef1234567890', 3, '2024-01-01']],
            columns: ['usertelemetryId', 'eventCount', 'lastSeen'],
            summary: '1 row',
            cached: false
        });

        await handlers.executeToolCall('lookup_user_telemetry_ids', {
            eventId: 'AL0000XYZ'
        });

        const kql: string = kusto.executeQuery.mock.calls[0][0];
        expect(kql).toContain('"AL0000XYZ"');
    });

    it('filters by aadTenantId when provided', async () => {
        const handlers = createHandlers();
        const kusto = (handlers.services as any).kusto;
        kusto.executeQuery = jest.fn().mockResolvedValue({
            type: 'table',
            rows: [['18c0f8e9-a967-47b0-abcd-ef1234567890', 1, '2024-01-01']],
            columns: ['usertelemetryId', 'eventCount', 'lastSeen'],
            summary: '1 row',
            cached: false
        });

        await handlers.executeToolCall('lookup_user_telemetry_ids', {
            aadTenantId: 'bc-customer-tenant'
        });

        const kql: string = kusto.executeQuery.mock.calls[0][0];
        expect(kql).toContain('bc-customer-tenant');
    });

    it('returns a friendly message when no telemetry events are found', async () => {
        const handlers = createHandlers();
        const kusto = (handlers.services as any).kusto;
        kusto.executeQuery = jest.fn().mockResolvedValue({
            type: 'table',
            rows: [],
            columns: [],
            summary: '0 rows',
            cached: false
        });

        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {
            eventId: 'NONEXISTENT'
        });

        expect(result.totalQueried).toBe(0);
        expect(result.message).toContain('NONEXISTENT');
        expect(mockLookupUsers).not.toHaveBeenCalled();
    });
});

// ─── Result structure ─────────────────────────────────────────────────────────

describe('lookup_user_telemetry_ids tool — result structure', () => {
    it('includes kqlFilter for each user', async () => {
        const handlers = createHandlers();

        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['18c0f8e9-a967-47b0-abcd-ef1234567890']
        });

        expect(result.users[0].kqlFilter).toBe(
            '| where tostring(customDimensions.usertelemetryId) == "18c0f8e9-a967-47b0-abcd-ef1234567890"'
        );
    });

    it('includes usage hints in the result', async () => {
        const handlers = createHandlers();

        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['18c0f8e9-a967-47b0-abcd-ef1234567890']
        });

        expect(result.usage).toBeDefined();
        expect(result.usage.summary).toContain('Resolved');
        expect(Array.isArray(result.usage.hints)).toBe(true);
    });

    it('surfaces resolvedVia field from lookup result', async () => {
        const handlers = createHandlers();

        const result = await handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['18c0f8e9-a967-47b0-abcd-ef1234567890']
        });

        expect(result.users[0].resolvedVia).toBe('bc_api');
    });
});

// ─── Tool configuration check ─────────────────────────────────────────────────

describe('lookup_user_telemetry_ids tool — configuration check', () => {
    it('throws when configuration is incomplete', async () => {
        const incompleteConfig: MCPConfig = { ...baseConfig, applicationInsightsAppId: '' };
        const services = initializeServices(incompleteConfig, true);
        const handlers = new ToolHandlers(
            incompleteConfig,
            services,
            true,
            ['BCTB_APP_INSIGHTS_ID is required']
        );

        await expect(handlers.executeToolCall('lookup_user_telemetry_ids', {
            usertelemetryIds: ['18c0f8e9-a967-47b0-abcd-ef1234567890']
        })).rejects.toThrow('Configuration incomplete');
    });
});
