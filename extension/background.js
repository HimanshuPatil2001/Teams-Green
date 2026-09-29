const ALARM_NAME      = 'teams-presence-refresh';
const WORK_CLIENT_ID  = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const DESKTOP_CLIENT_ID = '1fec8e78-bce4-4aaf-ab1b-5451cc387264';
const ALARM_PERIOD_MIN = 30;

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === ALARM_NAME) refreshPresence().catch(() => {});
});

chrome.runtime.onInstalled.addListener(() => ensureAlarm());
chrome.runtime.onStartup.addListener(() => ensureAlarm());

async function ensureAlarm() {
  const existing = await chrome.alarms.get(ALARM_NAME);
  if (!existing) {
    chrome.alarms.create(ALARM_NAME, { periodInMinutes: ALARM_PERIOD_MIN });
  }
}

// ── Main refresh loop ──────────────────────────────────────────────────────

async function refreshPresence() {
  const data = await chrome.storage.local.get([
    'sgsEnabled', 'graphToken', 'graphTokenExp', 'userEmail',
    'refreshToken', 'tenantId', 'localAccountId', 'clientId'
  ]);

  if (!data.sgsEnabled) return; 

  const now = Math.floor(Date.now() / 1000);
  const backendUrl = 'https://greenteam.himanshupatil.in';

  // Strategy 1: Stored Graph token is still valid
  if (data.graphToken && (data.graphTokenExp || 0) > now + 120) {
    const ok = await callSetPresence(data.graphToken);
    if (ok) { clearBadge(); return; }
  }

  // Strategy 2: acquireTokenSilent via open Teams tab
  const tab = await findTeamsTab();
  if (tab) {
    const result = await getTokenFromTab(tab.id).catch(() => ({}));
    if (result.token) {
      const tokenExp = result.exp || (now + 3600);
      await chrome.storage.local.set({
        graphToken: result.token,
        graphTokenExp: tokenExp,
      });

      // Push fresh token to our backend database
      pushGraphTokenToServer(backendUrl, result.token, tokenExp, data);
      
      const ok = await callSetPresence(result.token);
      if (ok) { clearBadge(); return; }
    }
  }

  // Strategy 3: Refresh token exchange (direct client-side refresh)
  if (data.refreshToken && data.tenantId && data.localAccountId) {
    const cid = data.clientId || WORK_CLIENT_ID;
    const token = await exchangeRefreshToken(
      data.refreshToken, data.tenantId, data.localAccountId, cid
    );
    if (token) {
      const tokenExp = now + (token.expires_in || 3600);
      await chrome.storage.local.set({
        graphToken: token.access_token,
        graphTokenExp: tokenExp,
      });
      if (token.refresh_token) {
        await chrome.storage.local.set({ refreshToken: token.refresh_token });
      }

      // Push updated tokens to backend
      pushGraphTokenToServer(backendUrl, token.access_token, tokenExp, {
        ...data,
        refreshToken: token.refresh_token || data.refreshToken
      });

      const ok = await callSetPresence(token.access_token);
      if (ok) { clearBadge(); return; }
    }
  }

  // All strategies failed — show warning badge
  setBadge('!');
}

// ── Token acquisition ──────────────────────────────────────────────────────

async function findTeamsTab() {
  const tabs = await chrome.tabs.query({ url: '*://teams.cloud.microsoft/*' });
  return tabs[0] || null;
}

async function getTokenFromTab(tabId) {
  const [result] = await chrome.scripting.executeScript({
    target: { tabId },
    world: 'MAIN',
    func: function() {
      function findMsalApp() {
        for (const k of ['__msal__', '__msalInstance', 'msalApp', '_msalApp', 'msalInstance', '__msalApp']) {
          try { if (window[k]?.acquireTokenSilent) return window[k]; } catch {}
        }
        for (const k of Object.getOwnPropertyNames(window)) {
          try {
            const v = window[k];
            if (v && typeof v === 'object' && typeof v.acquireTokenSilent === 'function'
                && typeof v.getAllAccounts === 'function') return v;
          } catch {}
        }
        return null;
      }
      const app = findMsalApp();
      if (!app) return { error: 'no_msal_app' };
      const account = app.getActiveAccount?.() || app.getAllAccounts?.()[0] || null;
      if (!account) return { error: 'no_account' };
      return app.acquireTokenSilent({
        account,
        scopes: ['https://graph.microsoft.com/Presence.ReadWrite'],
      }).then(r => ({
        token: r.accessToken,
        exp: r.expiresOn ? Math.floor(r.expiresOn.getTime() / 1000) : null,
      })).catch(e => ({ error: e.message?.slice(0, 100) }));
    },
  });
  return result?.result || {};
}

async function exchangeRefreshToken(refreshToken, tenantId, localAccountId, clientId) {
  try {
    const scope = clientId === DESKTOP_CLIENT_ID 
      ? 'https://api.spaces.skype.com/user_impersonation offline_access openid profile'
      : 'https://graph.microsoft.com/Presence.ReadWrite';

    const r = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded;charset=utf-8' },
        body: new URLSearchParams({
          client_id: clientId,
          scope: scope,
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          client_info: '1',
          'X-AnchorMailbox': `Oid:${localAccountId}@${tenantId}`,
        }),
      },
    );
    const j = await r.json();
    return j.access_token ? j : null;
  } catch { return null; }
}

// ── Graph API calls ────────────────────────────────────────────────────────

async function callSetPresence(token, skypeToken = null) {
  try {
    const sessionId = `sgs-bg-${Math.random().toString(36).slice(2, 10)}`;
    const headers = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'x-ms-client-user-agent': 'Teams-V2-Web',
      'x-ms-client-type': 'cdlworker',
      'x-ms-client-caller': 'TeamsClient:setPresenceAvailability',
      'x-ms-client-version': '1415/26080200616',
      'x-ms-endpoint-id': sessionId,
      'x-ms-session-id': sessionId,
      'behavioroverride': 'redirectAs404',
      'Referer': 'https://teams.cloud.microsoft/',
    };
    if (skypeToken) {
      headers['x-skypetoken'] = skypeToken;
    }

    // Attempt 1: Teams forceavailability API
    const r = await fetch('https://presence.teams.microsoft.com/v1/me/forceavailability/', {
      method: 'PUT',
      headers,
      body: JSON.stringify({ availability: 'Available' }),
    });

    if (r.ok || r.status === 204) return true;

    // Attempt 2: Graph API fallback
    const graphHeaders = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    };
    const rGraph = await fetch('https://graph.microsoft.com/v1.0/me/presence/setPresence', {
      method: 'POST',
      headers: graphHeaders,
      body: JSON.stringify({
        sessionId,
        availability: 'Available',
        activity: 'Available',
        expirationDuration: 'PT1H',
      }),
    });
    return rGraph.ok || rGraph.status === 204;
  } catch { return false; }
}

// ── Server-side token push ─────────────────────────────────────────────────

async function pushGraphTokenToServer(backendUrl, graphToken, graphTokenExp, data) {
  if (!data.localAccountId || !data.refreshToken || !graphToken) return;
  try {
    const payload = {
      userEmail: data.userEmail,
      tenantId: data.tenantId,
      microsoftGraphAccessToken: graphToken,
      microsoftGraphRefreshToken: data.refreshToken,
      tokenExpiryUtc: new Date(graphTokenExp * 1000).toISOString(),
      clientId: data.clientId,
      localAccountId: data.localAccountId
    };

    await fetch(`${backendUrl}/api/presence/extension-token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
  } catch (e) {
    console.error('[TeamsStayGreen SW] Background sync failed:', e);
  }
}

// ── Badge helpers ──────────────────────────────────────────────────────────

function clearBadge() {
  chrome.action.setBadgeText({ text: '' });
}

function setBadge(text) {
  chrome.action.setBadgeText({ text });
  chrome.action.setBadgeBackgroundColor({ color: '#e53e3e' });
}
