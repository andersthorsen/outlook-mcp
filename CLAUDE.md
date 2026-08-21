# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

- `npm install` - **ALWAYS run first** to install dependencies
- `npm start` - Start the MCP server
- `npm run auth-server` - Start the standalone OAuth server on port 3333 (optional fallback; the `authenticate` tool runs the flow in-process)
- `npm run test-mode` - Start the server in test mode with mock data
- `npm run inspect` - Use MCP Inspector to test the server interactively
- `npm test` - Run Jest tests
- `npx kill-port 3333` - Kill process using port 3333 if the OAuth callback server won't start

## Architecture Overview

This is a modular MCP (Model Context Protocol) server that provides Claude with access to Microsoft 365 services:
- **Outlook** - Email, calendar, folders, rules
- **OneDrive** - Files, folders, sharing
- **Power Automate** - Flows, environments, runs

### Core Structure
- `index.js` - Main entry point that combines all module tools and handles MCP protocol
- `config.js` - Centralized configuration (credentials, OAuth endpoints, scopes, field selections) — the single source of truth for auth settings
- `auth/auth-flow.js` - In-process OAuth flow: the `authenticate` tool starts a temporary callback server and returns the Microsoft sign-in URL
- `outlook-auth-server.js` - Standalone OAuth server (optional fallback, shares `config.js`)

### Modules
Each module exports tools and handlers:
- `auth/` - OAuth 2.0 authentication with token management (Graph + Flow)
- `calendar/` - Calendar operations (list, create, accept, decline, delete events)
- `email/` - Email management (list, search, read, send, mark as read)
- `folder/` - Folder operations (list, create, move)
- `rules/` - Email rules management
- `onedrive/` - OneDrive operations (list, search, download, upload, share, folder ops)
- `power-automate/` - Flow operations (list environments, list/run/toggle flows, run history)
- `utils/` - Shared utilities including Graph API client and OData helpers

### Key Components
- **Token Management**: Tokens stored in `~/.outlook-mcp-tokens.json` (both Graph and Flow tokens)
- **Graph API Client**: `utils/graph-api.js` handles Microsoft Graph API calls (Outlook, OneDrive)
- **Flow API Client**: `power-automate/flow-api.js` handles Power Automate API calls
- **Test Mode**: Mock data responses when `USE_TEST_MODE=true`
- **Modular Tools**: Each module exports tools array that gets combined in main server

## Authentication

### Graph API (Outlook + OneDrive)
1. Azure app registration required with permissions:
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
   - `Contacts.Read`
   - `User.Read`, `offline_access`
2. Use the `authenticate` tool — it starts an in-process callback server (port 3333, 5-minute window) and returns the Microsoft sign-in URL
3. Complete browser authentication; the MCP server handles the callback and saves the tokens
4. Tokens automatically stored and refreshed with the full scope set; a dead refresh token deletes the token file so `check-auth-status` reports honestly
5. `authenticate` with `force: true` re-authenticates (account switch, new scopes)

### Power Automate (Optional)
- Requires separate Flow API scope: `https://service.flow.microsoft.com//.default`
- Flow tokens stored alongside Graph tokens in same token file
- Only solution-aware flows accessible via API
- Only manual trigger flows can be triggered

## Configuration

### Environment Variables
- `MS_CLIENT_ID`/`MS_CLIENT_SECRET` and `OUTLOOK_CLIENT_ID`/`OUTLOOK_CLIENT_SECRET` are interchangeable; every process reads both (`MS_*` wins). `.env` is loaded by both the MCP server and the standalone auth server.
- `MS_TENANT_ID` (default `common`), `MS_AUTH_PORT` (default 3333), `MS_REDIRECT_URI`, `MS_SCOPES` override the defaults in `config.js`
- **Important**: Always use the client secret VALUE from Azure, not the Secret ID

### Config Constants
- `GRAPH_API_ENDPOINT`: `https://graph.microsoft.com/v1.0/`
- `FLOW_API_ENDPOINT`: `https://api.flow.microsoft.com`
- `ONEDRIVE_UPLOAD_THRESHOLD`: 4MB (files larger need chunked upload)
- Default page size: 25, max results: 50

### Common Setup Issues
1. **Missing dependencies**: Always run `npm install` first
2. **Wrong secret**: Use Azure secret VALUE, not ID (AADSTS7000215 error)
3. **Port conflicts**: the `authenticate` tool needs port 3333 free (or set `MS_AUTH_PORT`); use `npx kill-port 3333` if it is taken
4. **Container setups**: publish the callback port to the host so the browser redirect can reach it

## Test Mode

Set `USE_TEST_MODE=true` to use mock data instead of real API calls. Mock responses defined in:
- `utils/mock-data.js` - Graph API mocks
- `power-automate/flow-api.js` - Flow API mocks (inline)

## Error Handling

- Graph API auth failures: "UNAUTHORIZED" error
- Flow API auth failures: "FLOW_UNAUTHORIZED" error
- API errors include status codes and response details
- Token expiration triggers re-authentication flow
- Empty API responses handled gracefully
