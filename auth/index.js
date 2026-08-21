/**
 * Authentication module for Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const tokenStorage = require('./storage');
const { authTools } = require('./tools');

/**
 * Ensures the user is authenticated and returns an access token.
 * Automatically refreshes expired tokens using the refresh_token grant.
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated() {
  // Test tokens are fake and must never reach the real token refresh endpoint.
  if (config.USE_TEST_MODE) {
    const testToken = tokenManager.getAccessToken();
    if (!testToken) {
      throw new Error('Authentication required');
    }
    return testToken;
  }

  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated
};
