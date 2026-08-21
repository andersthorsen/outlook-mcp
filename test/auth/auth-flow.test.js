const http = require('http');

jest.mock('../../config', () => ({
  AUTH_CONFIG: {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    authEndpoint: 'https://login.example.test/authorize',
    tokenEndpoint: 'https://login.example.test/token',
    redirectUri: 'http://localhost:34567/auth/callback',
    authPort: 34567,
    scopes: ['offline_access', 'User.Read', 'Mail.Read'],
    tokenStorePath: '/mock/tokens.json'
  },
  USE_TEST_MODE: false
}));

const config = require('../../config');
const { AuthFlow } = require('../../auth/auth-flow');

function get(path) {
  // No connection pooling: each test runs its own short-lived server on the
  // same port, so a kept-alive socket would target an already-closed server.
  return new Promise((resolve, reject) => {
    http.get(`http://localhost:${config.AUTH_CONFIG.authPort}${path}`, { agent: false, headers: { Connection: 'close' } }, (res) => {
      let body = '';
      res.on('data', (chunk) => body += chunk);
      res.on('end', () => resolve({ statusCode: res.statusCode, body }));
    }).on('error', reject);
  });
}

function stateFromAuthUrl(authUrl) {
  return new URL(authUrl).searchParams.get('state');
}

describe('AuthFlow', () => {
  let tokenStorage;
  let flow;
  let consoleErrorSpy;

  beforeEach(() => {
    consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();
    tokenStorage = { exchangeCodeForTokens: jest.fn().mockResolvedValue({ access_token: 'new' }) };
    flow = new AuthFlow(tokenStorage);
  });

  afterEach(() => {
    flow.shutdown();
    consoleErrorSpy.mockRestore();
  });

  describe('start', () => {
    it('throws when client credentials are missing', async () => {
      const originalClientId = config.AUTH_CONFIG.clientId;
      config.AUTH_CONFIG.clientId = '';
      await expect(flow.start()).rejects.toThrow('client credentials are not configured');
      config.AUTH_CONFIG.clientId = originalClientId;
      expect(flow.isRunning()).toBe(false);
    });

    it('starts the callback server and returns a complete authorize URL', async () => {
      const result = await flow.start();

      expect(flow.isRunning()).toBe(true);
      expect(result.redirectUri).toBe(config.AUTH_CONFIG.redirectUri);
      expect(result.expiresInMinutes).toBe(5);

      const authUrl = new URL(result.authUrl);
      expect(result.authUrl.startsWith(config.AUTH_CONFIG.authEndpoint)).toBe(true);
      expect(authUrl.searchParams.get('client_id')).toBe('test-client-id');
      expect(authUrl.searchParams.get('response_type')).toBe('code');
      expect(authUrl.searchParams.get('redirect_uri')).toBe(config.AUTH_CONFIG.redirectUri);
      expect(authUrl.searchParams.get('scope')).toBe('offline_access User.Read Mail.Read');
      expect(authUrl.searchParams.get('state')).toMatch(/^[0-9a-f]{64}$/);
    });

    it('reuses the running server and keeps every pending state valid', async () => {
      const first = await flow.start();
      const second = await flow.start();

      expect(stateFromAuthUrl(first.authUrl)).not.toBe(stateFromAuthUrl(second.authUrl));
      expect(flow.pendingStates.size).toBe(2);
    });

    it('reports a helpful error when the port is taken', async () => {
      const blocker = http.createServer();
      await new Promise((resolve) => blocker.listen(config.AUTH_CONFIG.authPort, resolve));

      await expect(flow.start()).rejects.toThrow(`Port ${config.AUTH_CONFIG.authPort} is already in use`);

      await new Promise((resolve) => blocker.close(resolve));
    });
  });

  describe('callback handling', () => {
    it('returns 404 for paths other than the callback path', async () => {
      await flow.start();
      const res = await get('/somewhere-else');
      expect(res.statusCode).toBe(404);
    });

    it('rejects a callback with an unknown state', async () => {
      await flow.start();
      const res = await get('/auth/callback?code=abc&state=not-a-real-state');
      expect(res.statusCode).toBe(403);
      expect(tokenStorage.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    it('rejects a callback without a state', async () => {
      await flow.start();
      const res = await get('/auth/callback?code=abc');
      expect(res.statusCode).toBe(403);
      expect(tokenStorage.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    it('shows the provider error when the callback carries one', async () => {
      const { authUrl } = await flow.start();
      const state = stateFromAuthUrl(authUrl);
      const res = await get(`/auth/callback?error=access_denied&error_description=Nope&state=${state}`);
      expect(res.statusCode).toBe(400);
      expect(res.body).toContain('access_denied');
      expect(tokenStorage.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    it('rejects a callback without a code', async () => {
      const { authUrl } = await flow.start();
      const state = stateFromAuthUrl(authUrl);
      const res = await get(`/auth/callback?state=${state}`);
      expect(res.statusCode).toBe(400);
      expect(tokenStorage.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    it('exchanges the code and reports success', async () => {
      const { authUrl } = await flow.start();
      const state = stateFromAuthUrl(authUrl);
      const res = await get(`/auth/callback?code=auth-code-123&state=${state}`);

      expect(res.statusCode).toBe(200);
      expect(res.body).toContain('Authentication Successful');
      expect(tokenStorage.exchangeCodeForTokens).toHaveBeenCalledWith('auth-code-123');
    });

    it('consumes the state so it cannot be replayed', async () => {
      const { authUrl } = await flow.start();
      const state = stateFromAuthUrl(authUrl);
      await get(`/auth/callback?code=auth-code-123&state=${state}`);

      const replay = await get(`/auth/callback?code=another-code&state=${state}`);
      expect(replay.statusCode).toBe(403);
      expect(tokenStorage.exchangeCodeForTokens).toHaveBeenCalledTimes(1);
    });

    it('reports a failed token exchange', async () => {
      tokenStorage.exchangeCodeForTokens.mockRejectedValue(new Error('AADSTS7000215: bad secret'));
      const { authUrl } = await flow.start();
      const state = stateFromAuthUrl(authUrl);
      const res = await get(`/auth/callback?code=bad-code&state=${state}`);

      expect(res.statusCode).toBe(500);
      expect(res.body).toContain('AADSTS7000215');
    });
  });

  describe('shutdown', () => {
    it('stops the server and clears pending states', async () => {
      await flow.start();
      expect(flow.isRunning()).toBe(true);

      flow.shutdown();

      expect(flow.isRunning()).toBe(false);
      expect(flow.pendingStates.size).toBe(0);
      await expect(get('/auth/callback')).rejects.toThrow();
    });
  });
});
