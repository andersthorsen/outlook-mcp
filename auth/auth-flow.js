/**
 * In-process OAuth 2.0 authorization code flow.
 *
 * Started on demand by the `authenticate` tool: spins up a temporary HTTP
 * callback server on the configured auth port, hands back the Microsoft
 * authorize URL, exchanges the authorization code on callback, and shuts
 * the server down after success or a timeout.
 */
const http = require('http');
const crypto = require('crypto');
const querystring = require('querystring');
const config = require('../config');

const FLOW_TIMEOUT_MS = 5 * 60 * 1000;
const SHUTDOWN_DELAY_MS = 2000;

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderPage(title, heading, message, isError) {
  const color = isError ? '#d9534f' : '#5cb85c';
  const boxStyle = isError
    ? 'background-color: #f8d7da; border: 1px solid #f5c6cb;'
    : 'background-color: #d4edda; border: 1px solid #c3e6cb;';
  return `
    <html>
      <head>
        <title>${escapeHtml(title)}</title>
        <style>
          body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
          h1 { color: ${color}; }
          .box { ${boxStyle} padding: 15px; border-radius: 4px; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(heading)}</h1>
        <div class="box"><p>${message}</p></div>
        <p>You can close this window and return to your MCP client.</p>
      </body>
    </html>
  `;
}

class AuthFlow {
  constructor(tokenStorage) {
    this.tokenStorage = tokenStorage;
    this.server = null;
    this.pendingStates = new Map();
    this._timeout = null;
  }

  isRunning() {
    return this.server !== null;
  }

  /**
   * Starts (or reuses) the callback server and mints a fresh authorize URL.
   * @returns {Promise<{authUrl: string, redirectUri: string, expiresInMinutes: number}>}
   * @throws {Error} - If credentials are missing or the port cannot be bound
   */
  async start() {
    const { clientId, clientSecret, authEndpoint, redirectUri, scopes, authPort } = config.AUTH_CONFIG;

    if (!clientId || !clientSecret) {
      throw new Error(
        'Microsoft OAuth client credentials are not configured. ' +
        'Set MS_CLIENT_ID and MS_CLIENT_SECRET (or OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET) ' +
        'to the values from your Azure app registration.'
      );
    }

    if (!this.server) {
      await this._listen(authPort);
    }
    this._resetTimeout();

    const state = crypto.randomBytes(32).toString('hex');
    this.pendingStates.set(state, Date.now());

    const authUrl = `${authEndpoint}?${querystring.stringify({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      scope: scopes.join(' '),
      response_mode: 'query',
      state
    })}`;

    return { authUrl, redirectUri, expiresInMinutes: FLOW_TIMEOUT_MS / 60000 };
  }

  _listen(port) {
    return new Promise((resolve, reject) => {
      const server = http.createServer((req, res) => {
        this._handleRequest(req, res).catch((error) => {
          console.error('AuthFlow: unhandled error while serving callback:', error);
          res.writeHead(500, { 'Content-Type': 'text/plain' });
          res.end('Internal error');
        });
      });

      server.once('error', (error) => {
        if (error.code === 'EADDRINUSE') {
          reject(new Error(
            `Port ${port} is already in use, so the OAuth callback server could not start. ` +
            'Stop whatever is listening on that port (e.g. a standalone auth server or another MCP instance), ' +
            'or set MS_AUTH_PORT to a free port that is also registered as a redirect URI in Azure.'
          ));
        } else {
          reject(error);
        }
      });

      server.listen(port, () => {
        this.server = server;
        console.error(`AuthFlow: OAuth callback server listening on port ${port}`);
        resolve();
      });
    });
  }

  async _handleRequest(req, res) {
    const requestUrl = new URL(req.url, `http://localhost:${config.AUTH_CONFIG.authPort}`);
    const callbackPath = new URL(config.AUTH_CONFIG.redirectUri).pathname;

    if (requestUrl.pathname !== callbackPath) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not Found');
      return;
    }

    const query = Object.fromEntries(requestUrl.searchParams);

    if (!query.state || !this.pendingStates.has(query.state)) {
      console.error('AuthFlow: invalid or missing OAuth state parameter');
      res.writeHead(403, { 'Content-Type': 'text/html' });
      res.end(renderPage('Invalid State', 'Authentication Error',
        'Invalid or expired OAuth state parameter. Run the <code>authenticate</code> tool again to get a fresh sign-in link.', true));
      return;
    }
    this.pendingStates.delete(query.state);

    if (query.error) {
      console.error(`AuthFlow: authentication error: ${query.error}`);
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(renderPage('Authentication Error', 'Authentication Error',
        `<strong>Error:</strong> ${escapeHtml(query.error)}<br><strong>Description:</strong> ${escapeHtml(query.error_description || 'No description provided')}`, true));
      return;
    }

    if (!query.code) {
      console.error('AuthFlow: no authorization code provided in callback');
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(renderPage('Missing Authorization Code', 'Missing Authorization Code',
        'No authorization code was provided in the callback. Run the <code>authenticate</code> tool again.', true));
      return;
    }

    try {
      await this.tokenStorage.exchangeCodeForTokens(query.code);
      console.error('AuthFlow: token exchange successful');
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(renderPage('Authentication Successful', 'Authentication Successful!',
        'You have successfully authenticated with Microsoft Graph API. The tokens have been saved securely.', false));
      this._scheduleShutdown();
    } catch (error) {
      console.error(`AuthFlow: token exchange error: ${error.message}`);
      res.writeHead(500, { 'Content-Type': 'text/html' });
      res.end(renderPage('Token Exchange Error', 'Token Exchange Error', escapeHtml(error.message), true));
    }
  }

  _resetTimeout() {
    if (this._timeout) clearTimeout(this._timeout);
    this._timeout = setTimeout(() => {
      console.error('AuthFlow: timed out waiting for OAuth callback, shutting down');
      this.shutdown();
    }, FLOW_TIMEOUT_MS);
    this._timeout.unref();
  }

  _scheduleShutdown() {
    if (this._timeout) clearTimeout(this._timeout);
    this._timeout = setTimeout(() => this.shutdown(), SHUTDOWN_DELAY_MS);
    this._timeout.unref();
  }

  shutdown() {
    if (this._timeout) {
      clearTimeout(this._timeout);
      this._timeout = null;
    }
    this.pendingStates.clear();
    if (this.server) {
      this.server.close();
      this.server = null;
      console.error('AuthFlow: OAuth callback server stopped');
    }
  }
}

module.exports = { AuthFlow, FLOW_TIMEOUT_MS };
