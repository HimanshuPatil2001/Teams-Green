# Microsoft Teams Presence API - Manual Postman Testing Guide

This guide explains how to manually test Microsoft's Presence API using Postman and your browser session token before enabling the automated Hangfire scheduler.

---

## 📌 Prerequisites
1. **Google Chrome** or **Microsoft Edge** browser.
2. An active session on [Microsoft Teams Web](https://teams.cloud.microsoft) (or [teams.microsoft.com](https://teams.microsoft.com)).
3. **Postman** (or any API client / `curl`).

---

## 🔑 Step 1: Extract Active Bearer Token from Teams Web

1. Open your browser and navigate to **[https://teams.cloud.microsoft](https://teams.cloud.microsoft)** (or `https://teams.microsoft.com`).
2. Ensure you are signed in and your Teams dashboard is visible.
3. Open **Developer Tools** by pressing **`F12`** (or `Ctrl + Shift + I` on Windows / `Cmd + Option + I` on Mac).
4. Go to the **Network** tab in Developer Tools.
5. In the filter box, type **`presence`** or **`forceavailability`**.
6. Switch your status in Teams (e.g. change status to **Busy** or **Available**) to trigger presence network requests.
7. In the Network list, click on any matching request (e.g. `forceavailability`, `publishpresence`, or `me`).
8. In the right pane, select the **Headers** tab and scroll down to **Request Headers**.
9. Locate the **`Authorization`** header:
   ```text
   Bearer eyJ0eXAiOiJKV1QiLCJhbGci...
   ```
10. Copy the entire token string right after `Bearer ` (starts with `eyJ...`).

---

## 🚀 Step 2: Test Microsoft Presence API Directly in Postman

Create a new request in Postman to directly update your status via Microsoft's internal Teams presence endpoint:

### 1. HTTP Method & URL
* **Method**: `PUT`
* **URL**: `https://presence.teams.microsoft.com/v1/me/forceavailability/`

### 2. Request Headers
| Key | Value |
| :--- | :--- |
| `Authorization` | `Bearer YOUR_COPIED_TOKEN_HERE` |
| `Content-Type` | `application/json` |

### 3. Request Body
Select **raw** ➔ **JSON**:
```json
{
  "availability": "Available"
}
```

> **Available Status Options:**
> - `"Available"` (Green)
> - `"Busy"` (Red)
> - `"DoNotDisturb"` (Red)
> - `"BeRightBack"` (Yellow)
> - `"Away"` (Yellow)

### 4. Send & Verify
* Click **Send** in Postman.
* **Expected Response Status**: `200 OK` or `204 No Content`.
* **Verification**: Look at your Teams app / browser — your status badge should immediately reflect the chosen presence state (Green for Available).

---

## 🌐 Step 2 (Alternative): Test via Microsoft Graph Presence Endpoint

If your token was issued directly for Microsoft Graph (`https://graph.microsoft.com`):

### 1. HTTP Method & URL
* **Method**: `POST`
* **URL**: `https://graph.microsoft.com/v1.0/me/presence/setPresence`

### 2. Request Headers
| Key | Value |
| :--- | :--- |
| `Authorization` | `Bearer YOUR_COPIED_TOKEN_HERE` |
| `Content-Type` | `application/json` |

### 3. Request Body (raw JSON)
```json
{
  "sessionId": "d3b07384-d113-495f-a558-18e974e359a3",
  "availability": "Available",
  "activity": "Available",
  "expirationDuration": "PT1H"
}
```

---

## ⚙️ Step 3: Test via Your Local Backend API

Once you have verified that the token works directly against Microsoft, test your backend API to register the session and trigger an on-demand server ping.

> Make sure your backend API is running (`dotnet run` or via Visual Studio / Docker). Default local port is usually `http://localhost:5000` or `http://localhost:8091`.

### 1. Register Session with Token
* **Method**: `POST`
* **URL**: `http://localhost:5000/api/presence/extension-token`
* **Headers**: `Content-Type: application/json`
* **Body (raw JSON)**:
  ```json
  {
    "accessToken": "YOUR_COPIED_TOKEN_HERE",
    "targetPresenceState": "Available"
  }
  ```
* **Expected Response (`200 OK`)**:
  ```json
  {
    "id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
    "tenantId": "...",
    "userEmail": "your.email@domain.com",
    "isActive": true,
    "targetPresenceState": "Available"
  }
  ```
* Copy the `id` (GUID) from the response.

### 2. Trigger On-Demand Backend Keep-Alive Ping
* **Method**: `POST`
* **URL**: `http://localhost:5000/api/presence/ping/{YOUR_SESSION_ID}`
* **Expected Response (`200 OK`)**:
  ```json
  {
    "message": "Ping executed.",
    "lastResponseCode": 200,
    "lastResponseBody": "HTTP Status 200"
  }
  ```

### 3. Check Session Logs
* **Method**: `GET`
* **URL**: `http://localhost:5000/api/presence/logs/{YOUR_SESSION_ID}`
* **Expected Response (`200 OK`)**: Returns an array of recent ping logs with timestamps and HTTP status codes.

---

## ⚠️ Troubleshooting & Common Issues

| Issue / Error | Cause | Resolution |
| :--- | :--- | :--- |
| **`401 Unauthorized`** | Token expired or copied improperly. | Microsoft browser access tokens usually expire in 60 minutes. Grab a fresh token from Developer Tools. |
| **`403 Forbidden`** | Scope mismatch or tenant conditional access restrictions. | Ensure you copied the token from an active `presence` request on `teams.cloud.microsoft`. |
| **Backend connection refused** | Backend API is not running. | Run `dotnet run` inside the project folder or verify your Docker container port mapping. |
