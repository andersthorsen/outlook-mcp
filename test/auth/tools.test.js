const mockFlowStart = jest.fn();

jest.mock('../../config', () => ({
  SERVER_VERSION: '2.0.0',
  USE_TEST_MODE: false,
  AUTH_CONFIG: {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tokenStorePath: '/mock/tokens.json'
  }
}));

jest.mock('../../auth/storage', () => ({
  getValidAccessToken: jest.fn(),
  getExpiryTime: jest.fn(),
  getTokens: jest.fn(),
  isTokenExpired: jest.fn(),
  clearTokens: jest.fn()
}));

jest.mock('../../auth/token-manager', () => ({
  createTestTokens: jest.fn(),
  loadTokenCache: jest.fn()
}));

jest.mock('../../auth/auth-flow', () => ({
  AuthFlow: jest.fn().mockImplementation(() => ({ start: mockFlowStart }))
}));

const config = require('../../config');
const tokenStorage = require('../../auth/storage');
const tokenManager = require('../../auth/token-manager');
const { handleAuthenticate, handleCheckAuthStatus } = require('../../auth/tools');

function textOf(response) {
  return response.content[0].text;
}

describe('auth tools', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    config.USE_TEST_MODE = false;
  });

  describe('handleAuthenticate', () => {
    const flowResult = {
      authUrl: 'https://login.example.test/authorize?client_id=test-client-id&state=abc',
      redirectUri: 'http://localhost:3333/auth/callback',
      expiresInMinutes: 5
    };

    it('creates test tokens in test mode', async () => {
      config.USE_TEST_MODE = true;
      const response = await handleAuthenticate({});
      expect(tokenManager.createTestTokens).toHaveBeenCalled();
      expect(textOf(response)).toContain('test mode');
    });

    it('short-circuits when already authenticated', async () => {
      const expiresAt = Date.now() + 3600000;
      tokenStorage.getValidAccessToken.mockResolvedValue('valid-token');
      tokenStorage.getExpiryTime.mockReturnValue(expiresAt);

      const response = await handleAuthenticate({});

      expect(textOf(response)).toContain('Already authenticated');
      expect(textOf(response)).toContain(new Date(expiresAt).toISOString());
      expect(mockFlowStart).not.toHaveBeenCalled();
    });

    it('starts the OAuth flow and returns the sign-in URL when not authenticated', async () => {
      tokenStorage.getValidAccessToken.mockResolvedValue(null);
      mockFlowStart.mockResolvedValue(flowResult);

      const response = await handleAuthenticate({});

      expect(mockFlowStart).toHaveBeenCalled();
      expect(tokenStorage.clearTokens).not.toHaveBeenCalled();
      expect(textOf(response)).toContain(flowResult.authUrl);
      expect(textOf(response)).toContain(flowResult.redirectUri);
      expect(textOf(response)).toContain('5 minutes');
    });

    it('clears tokens and re-authenticates when force is set', async () => {
      mockFlowStart.mockResolvedValue(flowResult);

      const response = await handleAuthenticate({ force: true });

      expect(tokenStorage.clearTokens).toHaveBeenCalled();
      expect(tokenStorage.getValidAccessToken).not.toHaveBeenCalled();
      expect(textOf(response)).toContain(flowResult.authUrl);
    });

    it('reports a failure to start the flow', async () => {
      tokenStorage.getValidAccessToken.mockResolvedValue(null);
      mockFlowStart.mockRejectedValue(new Error('Port 3333 is already in use'));

      const response = await handleAuthenticate({});

      expect(textOf(response)).toContain('Could not start the authentication flow');
      expect(textOf(response)).toContain('Port 3333 is already in use');
    });
  });

  describe('handleCheckAuthStatus', () => {
    it('reports test-mode status from the test token cache', async () => {
      config.USE_TEST_MODE = true;
      tokenManager.loadTokenCache.mockReturnValue({ access_token: 'test_access_token_1' });
      expect(textOf(await handleCheckAuthStatus())).toBe('Authenticated and ready (test mode)');

      tokenManager.loadTokenCache.mockReturnValue(null);
      expect(textOf(await handleCheckAuthStatus())).toBe('Not authenticated (test mode)');
    });

    it('reports not authenticated when there are no tokens', async () => {
      tokenStorage.getTokens.mockResolvedValue(null);
      const response = await handleCheckAuthStatus();
      expect(textOf(response)).toContain('Not authenticated');
      expect(textOf(response)).toContain('authenticate');
    });

    it('reports expiry and scopes for a valid token', async () => {
      const expiresAt = Date.now() + 3600000;
      tokenStorage.getTokens.mockResolvedValue({ access_token: 'x', scope: 'User.Read Mail.Read' });
      tokenStorage.isTokenExpired.mockReturnValue(false);
      tokenStorage.getExpiryTime.mockReturnValue(expiresAt);

      const text = textOf(await handleCheckAuthStatus());
      expect(text).toContain('Authenticated and ready');
      expect(text).toContain(new Date(expiresAt).toISOString());
      expect(text).toContain('User.Read Mail.Read');
    });

    it('reports success when an expired token refreshes automatically', async () => {
      const expiresAt = Date.now() + 3600000;
      tokenStorage.getTokens.mockResolvedValue({ access_token: 'x', refresh_token: 'r' });
      tokenStorage.isTokenExpired.mockReturnValue(true);
      tokenStorage.getValidAccessToken.mockResolvedValue('refreshed');
      tokenStorage.getExpiryTime.mockReturnValue(expiresAt);

      const text = textOf(await handleCheckAuthStatus());
      expect(text).toContain('refreshed automatically');
      expect(text).toContain(new Date(expiresAt).toISOString());
    });

    it('asks the user to re-authenticate when refresh fails', async () => {
      tokenStorage.getTokens.mockResolvedValue({ access_token: 'x', refresh_token: 'dead' });
      tokenStorage.isTokenExpired.mockReturnValue(true);
      tokenStorage.getValidAccessToken.mockResolvedValue(null);

      const text = textOf(await handleCheckAuthStatus());
      expect(text).toContain('could not be refreshed');
      expect(text).toContain('authenticate');
    });
  });
});
