/**
 * Shared TokenStorage instance used by the MCP server process.
 * Keeping a single instance means the in-memory token cache stays
 * consistent between tool calls and the in-process auth flow.
 */
const TokenStorage = require('./token-storage');

module.exports = new TokenStorage();
