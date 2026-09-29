Here is a highly detailed, professional-grade **"Antigravity IDE Prompt"** designed to be pasted directly into an AI-powered IDE assistant (such as Claude Code, Cursor, or GitHub Copilot) inside your newly created .NET 8.0 Web API directory.

It is engineered to build a clean, production-ready, multitenant-capable (SaaS-proof) architecture using **ASP.NET Core Controllers**, **Entity Framework Core**, **Hangfire**, and **Docker**.

---

### 📋 How to use this prompt:

1. Create your .NET 8.0 Web API project via CLI or Visual Studio (ensuring you select Linux Docker support, Controllers, and Swagger/OpenAPI).


2. Open the project folder in your IDE.
3. Copy the entire prompt below and paste it into your AI assistant chat window (Composer, Agent, or Chat interface) and watch it generate the code.

---

# Role: Senior .NET Software Architect & SaaS Engineer

# Objective: Build a robust, SaaS-ready "Teams Status Keep-Alive" engine inside this .NET 8.0 Web API project.

We are building a production-grade backend engine that programmatically keeps Microsoft Teams presence indicators active via the official Microsoft Graph API. The system must support single-user local debugging immediately, but use a database-driven multi-tenant (SaaS-ready) structural layout.

## Tech Stack Requirements

* Target: .NET 8.0 Web API
* Architecture: Controller-based Routing (with Swagger/OpenAPI enabled)
* Database: EF Core with SQLite (for easy local setup, structurally ready to swap to PostgreSQL/SQL Server)
* Background Engine: Hangfire (for job scheduling & dashboard views)
* External Integrations: Microsoft Graph API `/me/presence/setPresence`

* Deployment: Optimized Dockerfile with a Linux alpine-based multi-stage container build

---

## Technical Specifications & Architecture

### 1. Database & SaaS-Ready Entities (EF Core)

Create a database context (`KeepAliveDbContext`) with the following entities:

* **TenantSession**: Tracks user authentication details. Must include:
* `Id` (Guid, primary key)
* `TenantId` (String, mapped to corporate AD directory or Stripe customer ID)
* `UserEmail` (String, unique identifier for logging)
* `MicrosoftGraphAccessToken` (String)
* `MicrosoftGraphRefreshToken` (String)
* `TokenExpiryUtc` (DateTime)
* `IsActive` (Boolean)
* `TargetPresenceState` (String - default "Available")




* **PresenceLog**: Log table tracking each keep-alive ping. Includes `Id`, `TenantSessionId`, `TimestampUtc`, `HttpStatusCode`, and `ResponseBody`.

### 2. Microsoft Graph Presence Service

Write a transient/scoped service (`IMSGraphPresenceService`) that:

* Refreshes expired Graph tokens securely using Graph OAuth endpoints.
* Constructs and posts the exact JSON body to `[https://graph.microsoft.com/v1.0/me/presence/setPresence](https://graph.microsoft.com/v1.0/me/presence/setPresence)`.


* Payload structure must match Microsoft's API schema:json
{
"sessionId": "60d8d816-7a65-4bde-90c7-70cc0d0e9b60",
"availability": "Available",
"activity": "Available",
"expirationDuration": "PT1H"
}
```
*(Note: Expiration duration should be formatted to ISO 8601 duration "PT1H" [cite: 4]).*

```


* Automatically records the API response status and message inside `PresenceLog`.

### 3. Hangfire Setup & Background Work Dispatcher

* Configure Hangfire in `Program.cs` to use the database storage provider.
* Set up the Hangfire Dashboard to be accessible at `/hangfire`.
* **SaaS-Proof Security:** Add a custom `IDashboardAuthorizationFilter` called `HangfireDashboardAuthorizationFilter`. For single-user testing, let it bypass checks locally, but structure it to validate admin credentials or JWT values if hosted in production.
* Register a **Recurring Job** that executes every **4 minutes**. This job must:


1. Fetch all `IsActive = true` sessions from the DB.
2. For each active session, run the keep-alive Graph API loop.


3. Ensure exception-handling and retry logic are in place so one failing session does not crash other jobs.



### 4. Controller Layer (Traditional Controllers)

Implement a `PresenceController` inheriting from `ControllerBase`:

* `POST api/presence/session`: Add or update a tenant session (providing access token, refresh token, expiry, and email).
* `POST api/presence/toggle/{id}`: Enable or disable a session.
* `POST api/presence/ping/{id}`: Manual override endpoint to immediately trigger a keep-alive call for debugging.
* `GET api/presence/logs/{id}`: Query the last 50 execution logs for a session to review historical activity.
* Document all routes using OpenAPI/Swagger annotations.

---

## Implementation Steps

Please implement or modify the following files in this project:

1. **NuGet Packages to Install:**
* `Hangfire.AspNetCore`
* `Hangfire.Storage.SQLite` (or use SqlServer/PostgreSQL based on target DbContext setup)
* `Microsoft.EntityFrameworkCore.Sqlite`
* `Microsoft.EntityFrameworkCore.Design`


2. **File Modifications:**
* Create `Entities/TenantSession.cs` and `Entities/PresenceLog.cs`.
* Create `Data/KeepAliveDbContext.cs` and handle migrations securely.
* Create `Services/MSGraphPresenceService.cs` and interface.
* Create `Hangfire/HangfireDashboardAuthorizationFilter.cs`.
* Create `Controllers/PresenceController.cs`.
* Modify `Program.cs` to integrate controllers, EF Core SQLite DB, Swagger, and register Hangfire background server/dashboard pipeline.
* Output a production-grade multi-stage **Dockerfile** targeting Alpine Linux for rapid container build and small image footprint.



Ensure the code is robust, asynchronous (using async/await throughout), handles API exceptions gracefully, and prints structured console logs using `ILogger`.

---

## 💡 SaaS Flow Architecture & Developer Knowledge Base

### 1. Two-Way SaaS Authentication Flow
To support all user types (especially in strict corporate environments), the SaaS utilizes two separate flows:
* **The OAuth Web Flow (Primary / Preferred)**:
  * **Target**: Personal accounts and cooperative Work/School accounts.
  * **Action**: User logs in via the dashboard using Microsoft OAuth screen requesting `offline_access` and `https://graph.microsoft.com/.default` scopes.
  * **Benefit**: The backend receives a `RefreshToken` to renew access tokens automatically. This enables **24/7 hands-off keep-alive** even when the user's laptop is closed.
* **The Chrome Extension Flow (Secondary / Fallback)**:
  * **Target**: Work/School accounts where the organization's Tenant Admin has blocked third-party Azure AD app consent.
  * **Action**: The Chrome Extension monitors background traffic in the browser, captures the Microsoft Graph token, and syncs it to `/api/presence/extension-token`.
  * **Benefit**: Bypasses Azure App consent blocks. However, since no refresh token is captured in this flow, the keep-alive only works in ~4-hour intervals while the browser Teams page is open.

### 2. Personal Accounts (`teams.live.com`) vs. Work Accounts (`teams.microsoft.com`)
* **Personal Accounts**:
  * Run on Microsoft's consumer/Skype infrastructure.
  * The browser client calls consumer APIs like `https://teams.live.com/ups/global/v1/me/forceavailability/` using JWE (5-part) encrypted and Skype-specific tokens.
  * **Note**: Personal accounts in the browser **do not** make calls to `graph.microsoft.com`. The Chrome Extension will not capture standard Graph tokens on `teams.live.com`.
  * **Solution**: Personal account users should always use the **OAuth Web Flow** on the dashboard.
* **Work/School Accounts**:
  * Run on enterprise infrastructure.
  * The browser client frequently calls `https://graph.microsoft.com/...` for calendars, OneDrive, profile details, etc.
  * **Solution**: Fully compatible with the Chrome Extension flow.

### 3. Developer Testing & Troubleshooting

#### How to Extract a Live Token from Developer Tools (F12)
1. Open the Teams Web App and open Developer Tools (`F12` -> **Network** tab).
2. Set the Filter box to: `graph.microsoft.com`.
3. In the sub-filters row, ensure **`Fetch/XHR`** (or `All`) is selected (do not select `JS`).
4. Click on your profile picture, files, or OneDrive tab in Teams to trigger a Graph API request.
5. Select a request (e.g. `/v1/me`), open the **Headers** tab, scroll to **Request Headers**, and copy the **`Authorization`** token (starts with `Bearer ey...` and contains exactly 2 dots).

#### Sandbox testing of the Chrome Extension Flow
Since personal accounts do not make Graph calls in the browser, sign up for a free **[Microsoft 365 Developer Tenant](https://developer.microsoft.com/en-us/microsoft-365/dev-program)** to get a sandbox environment. This gives you a test enterprise tenant and developer accounts (e.g., `admin@companyname.onmicrosoft.com`) to test the Chrome Extension flow end-to-end.

#### Mock Tokens for Local API Testing
When testing API endpoints (like `/api/presence/extension-token`) with mock JWT payloads:
* Ensure mock tokens contain exactly **2 dots (3 parts)** (e.g. `mockHeader.mockPayload.mockSignature`).
* If you use a token with only 1 dot, Microsoft Graph API calls will return a `401 Unauthorized` with the error `IDX14120: JWT is not well formed, there is only one dot (.).` before verifying the signature.

#### Troubleshooting: "Refresh token not found in MSAL cache. Reload Teams and try again"
* **Symptom:** When using browser-based token extraction extensions, you may encounter an error message: `"Refresh token not found in MSAL cache. Reload Teams and try again."`
* **Root Cause:** Personal/Consumer accounts (`teams.live.com` / `MsaLife`) use Skype-based token infrastructure (`skypeToken` and `api.fl.spaces.skype.com`) rather than standard Microsoft Entra ID (Azure AD) MSAL-cached tokens for status updates. Therefore, no standard Graph refresh token is stored in the browser's MSAL cache.
* **Resolution:** 
  1. Personal account users must bypass browser-based extension flows and instead use the **OAuth Web Flow** in the web app dashboard, consenting to the `Presence.ReadWrite` scope.
  2. If testing browser extension token extraction flows is required, you must use a Work/School account or provision a Microsoft 365 Developer Tenant sandbox account.