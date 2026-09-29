# Teams Stay Green API - Testing Guide

Use these pre-built, related `curl` commands to test your API. The values are aligned and compatible, including a valid mock Microsoft Graph JWT access token that the API can successfully parse.

### 🔌 Docker Port Mappings Detected:
- **HTTP Endpoint**: **`http://localhost:32768`** (Maps to container port `8080`) -> **Use this to avoid SSL certificate issues!**
- **HTTPS Endpoint**: **`https://localhost:32769`** (Maps to container port `8081`)

---

## 🔑 Constant Test Values Used:
- **User Email**: `user@example.com`
- **Tenant ID**: `common`
- **Session GUID**: `d3b07384-d113-495f-a558-18e974e359a3`
- **Presence State**: `Available`
- **Mock Token**: A valid base64-encoded mock token containing `{ "upn": "user@example.com", "tid": "common", "exp": 1799999999 }` (appended with a mock signature to form a valid 3-part JWS structure)

---

## 1. Get Public Configuration
Checks if the API is running and returns configuration details.
```bash
curl -X GET "http://localhost:32768/api/presence/config" \
     -H "accept: application/json"
```

## 2. Sync Token (Extension simulation)
Simulates the Chrome Extension capturing a token from Teams and sending it to the backend. This parses the email (`user@example.com`) directly from the token and registers/updates the session.
```bash
curl -X POST "http://localhost:32768/api/presence/extension-token" \
     -H "accept: application/json" \
     -H "Content-Type: application/json" \
     -d '{
       "accessToken": "mockHeader.eyJ1cG4iOiJ1c2VyQGV4YW1wbGUuY29tIiwidGlkIjoiY29tbW9uIiwiZXhwIjoxNzk5OTk5OTk5fQ==.mockSignature",
       "targetPresenceState": "Available"
     }'
```

## 3. Retrieve All Active Sessions
Queries all sessions in the DB to see if the session was created successfully.
```bash
curl -X GET "http://localhost:32768/api/presence/sessions" \
     -H "accept: application/json"
```

## 4. Manual Create/Update Session
Registers or updates a session manually with custom token details.
```bash
curl -X POST "http://localhost:32768/api/presence/session" \
     -H "accept: application/json" \
     -H "Content-Type: application/json" \
     -d '{
       "tenantId": "common",
       "userEmail": "user@example.com",
       "microsoftGraphAccessToken": "mockHeader.eyJ1cG4iOiJ1c2VyQGV4YW1wbGUuY29tIiwidGlkIjoiY29tbW9uIiwiZXhwIjoxNzk5OTk5OTk5fQ==.mockSignature",
       "microsoftGraphRefreshToken": "mock_refresh_token_12345",
       "tokenExpiryUtc": "2027-01-01T00:00:00Z",
       "targetPresenceState": "Available"
     }'
```

## 5. Toggle Session Active Status
Toggles the active state (`IsActive`) of the registered session.
> [!IMPORTANT]
> Make sure to replace `d3b07384-d113-495f-a558-18e974e359a3` with the actual **GUID** returned by the `sessions` endpoint if it differs.
```bash
curl -X POST "http://localhost:32768/api/presence/toggle/d3b07384-d113-495f-a558-18e974e359a3" \
     -H "accept: application/json"
```

## 6. Manual Keep-Alive Ping (Override)
Forces the background worker logic to run immediately for this session.
```bash
curl -X POST "http://localhost:32768/api/presence/ping/d3b07384-d113-495f-a558-18e974e359a3" \
     -H "accept: application/json"
```

## 7. Get Session Execution Logs
Gets the execution logs and response history of the pings for this session.
```bash
curl -X GET "http://localhost:32768/api/presence/logs/d3b07384-d113-495f-a558-18e974e359a3" \
     -H "accept: application/json"
```

