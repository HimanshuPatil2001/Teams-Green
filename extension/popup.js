const CONSUMER_CLIENT_ID = '4b3e8f46-56d3-427f-b1e2-d239b2ea6bca';
const WORK_CLIENT_ID     = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const CONSUMER_TENANT    = '9188040d-6c67-4c5b-b112-36a304b66dad';
const MSA_TENANT         = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a';

const DESKTOP_CLIENT_ID  = '1fec8e78-bce4-4aaf-ab1b-5451cc387264';
const NATIVECLIENT_REDIRECT = 'https://login.microsoftonline.com/common/oauth2/nativeclient';
const DESKTOP_SCOPE      = 'https://api.spaces.skype.com/user_impersonation offline_access openid profile';

function sgtB64Url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// Silently obtain a Desktop client refresh token for a work account by running a prompt=none auth flow
async function getDesktopRTViaAuthCode(tenantId, backendUrl) {
  if (!tenantId) return null;
  let verifier, authUrl;
  try {
    verifier = sgtB64Url(crypto.getRandomValues(new Uint8Array(48)));
    const challenge = sgtB64Url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?` + new URLSearchParams({
      client_id: DESKTOP_CLIENT_ID,
      response_type: 'code',
      redirect_uri: NATIVECLIENT_REDIRECT,
      scope: DESKTOP_SCOPE,
      response_mode: 'query',
      code_challenge: challenge,
      code_challenge_method: 'S256',
      prompt: 'none',
    }).toString();
  } catch (e) {
    console.error('[TeamsStayGreen] Failed to generate OAuth URL:', e);
    return null;
  }

  const code = await new Promise((resolve) => {
    let settled = false, createdTabId = null;
    const finish = (val) => {
      if (settled) return; settled = true;
      chrome.tabs.onUpdated.removeListener(listener);
      clearTimeout(timer);
      if (createdTabId != null) chrome.tabs.remove(createdTabId).catch(() => {});
      resolve(val);
    };
    const listener = (tabId, changeInfo) => {
      if (tabId !== createdTabId || !changeInfo.url) return;
      if (!changeInfo.url.startsWith(NATIVECLIENT_REDIRECT)) return;
      try {
        const u = new URL(changeInfo.url);
        finish(u.searchParams.get('code') || null);
      } catch { finish(null); }
    };
    chrome.tabs.onUpdated.addListener(listener);
    const timer = setTimeout(() => finish(null), 15000);
    chrome.tabs.create({ url: authUrl, active: false }, (tab) => { createdTabId = tab?.id ?? null; });
  });
  if (!code) return null;

  try {
    const r = await fetch(`${backendUrl}/api/presence/redeem-code`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ code, codeVerifier: verifier, tenantId }),
    });
    const j = await r.json();
    return j?.refreshToken || null;
  } catch (e) {
    console.error('[TeamsStayGreen] Failed to redeem code via backend:', e);
    return null;
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  const targetStateSelect = document.getElementById('targetState');
  const backendUrlInput = document.getElementById('backendUrl');
  const syncStatusEl = document.getElementById('syncStatus');
  const syncTimeEl = document.getElementById('syncTime');
  const saveConfigBtn = document.getElementById('saveConfig');
  const openTeamsBtn = document.getElementById('openTeams');

  const backendUrl = backendUrlInput.value || 'https://greenteam.himanshupatil.in';

  // Load saved configuration
  chrome.storage.local.get([
    'targetPresenceState',
    'lastSyncStatus',
    'lastSyncTime'
  ], function(data) {
    if (data.targetPresenceState) {
      targetStateSelect.value = data.targetPresenceState;
    }
    updateStatusEl(data.lastSyncStatus || 'Waiting...');
    if (data.lastSyncTime) {
      syncTimeEl.textContent = data.lastSyncTime;
    }
  });

  // Open Teams Web
  openTeamsBtn.addEventListener('click', () => {
    chrome.tabs.create({ url: 'https://teams.cloud.microsoft/' });
  });

  // Connect & Sync
  saveConfigBtn.addEventListener('click', async () => {
    saveConfigBtn.disabled = true;
    saveConfigBtn.textContent = 'Syncing...';
    updateStatusEl('Reading tokens...');

    try {
      // Find active Teams tab
      const allTeamsTabs = await chrome.tabs.query({
        url: ['https://teams.live.com/*', 'https://teams.cloud.microsoft/*'],
      });

      if (allTeamsTabs.length === 0) {
        throw new Error('Please open Microsoft Teams Web first.');
      }

      const targetTab = allTeamsTabs[0];
      let targetHost;
      try { targetHost = new URL(targetTab.url).hostname; } catch { targetHost = 'teams.cloud.microsoft'; }
      const isTargetWork = targetHost === 'teams.cloud.microsoft';

      // Inject extraction script into Teams Web context
      const [scriptResult] = await chrome.scripting.executeScript({
        target: { tabId: targetTab.id },
        func: extractFromMSAL,
        args: [CONSUMER_CLIENT_ID, WORK_CLIENT_ID, CONSUMER_TENANT, MSA_TENANT],
      });
      const found = scriptResult?.result || {};

      if (found.error) throw new Error(found.error);
      if (!found.refresh_token) {
        throw new Error('No login session found. Please refresh Teams Web page.');
      }

      const isWork = (found.account_type === 'work') || isTargetWork;
      let finalRefreshToken = found.refresh_token;
      let clientId = isWork ? WORK_CLIENT_ID : CONSUMER_CLIENT_ID;

      // Swap Work RT for Desktop Client RT silently (allows 90 days offline keep-alive)
      if (isWork) {
        updateStatusEl('Securing long-lived session...');
        const desktopRt = await getDesktopRTViaAuthCode(found.work_tenant || found.tenant_id, backendUrl);
        if (desktopRt) {
          finalRefreshToken = desktopRt;
          clientId = DESKTOP_CLIENT_ID;
        }
      }

      // Query active presence access token scoped to Graph
      let graphToken = null;
      let graphTokenExp = null;
      if (isWork) {
        const [gResult] = await chrome.scripting.executeScript({
          target: { tabId: targetTab.id },
          world: 'MAIN',
          func: function() {
            function findMsalApp() {
              for (const k of ['__msal__', '__msalInstance', 'msalApp', '_msalApp', 'msalInstance', '__msalApp']) {
                try { if (window[k]?.acquireTokenSilent) return window[k]; } catch {}
              }
              for (const k of Object.getOwnPropertyNames(window)) {
                try {
                  const v = window[k];
                  if (v && typeof v === 'object' && typeof v.acquireTokenSilent === 'function') return v;
                } catch {}
              }
              return null;
            }
            const app = findMsalApp();
            if (!app) return null;
            const account = app.getActiveAccount?.() || app.getAllAccounts?.()[0] || null;
            if (!account) return null;
            return app.acquireTokenSilent({
              account,
              scopes: ['https://graph.microsoft.com/Presence.ReadWrite'],
            }).then(r => ({
              token: r.accessToken,
              exp: r.expiresOn ? Math.floor(r.expiresOn.getTime() / 1000) : null,
            })).catch(() => null);
          }
        });
        if (gResult?.result?.token) {
          graphToken = gResult.result.token;
          graphTokenExp = gResult.result.exp;
        }
      }

      // If we couldn't get a Graph-specific token from the cache (e.g. personal account),
      // we fallback to sending the captured access token
      const finalAccessToken = graphToken || found.presence_access_token || found.spaces_access_token;
      
      const payload = {
        userEmail: found.username || found.userEmail || `teams_${found.local_account_id}@staygreen.com`,
        tenantId: found.work_tenant || found.tenant_id || 'common',
        microsoftGraphAccessToken: finalAccessToken || '',
        microsoftGraphRefreshToken: finalRefreshToken || '',
        tokenExpiryUtc: graphTokenExp ? new Date(graphTokenExp * 1000).toISOString() : new Date(Date.now() + 3600 * 1000).toISOString(),
        targetPresenceState: targetStateSelect.value,
        clientId: clientId,
        localAccountId: found.local_account_id
      };

      updateStatusEl('Syncing to server...');
      
      // POST to our C# Backend API
      const response = await fetch(`${backendUrl}/api/presence/extension-token`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server failed to sync: ${errorText || response.statusText}`);
      }

      // Save sync state
      const timeString = new Date().toLocaleTimeString();
      chrome.storage.local.set({
        targetPresenceState: targetStateSelect.value,
        lastSyncStatus: 'Connected',
        lastSyncTime: timeString,
        sgsEnabled: true,
        refreshToken: finalRefreshToken,
        tenantId: payload.tenantId,
        localAccountId: payload.localAccountId,
        clientId: clientId,
        userEmail: payload.userEmail
      });

      updateStatusEl('Connected');
      syncTimeEl.textContent = timeString;
      alert('Successfully synced with GreenTeam Server!');

    } catch (e) {
      console.error(e);
      updateStatusEl('Failed');
      alert(`Sync Error: ${e.message}`);
    } finally {
      saveConfigBtn.disabled = false;
      saveConfigBtn.textContent = 'Save Configuration';
    }
  });

  function updateStatusEl(status) {
    syncStatusEl.textContent = status;
    syncStatusEl.className = 'status-badge';
    if (status === 'Connected') {
      syncStatusEl.classList.add('connected');
    } else if (status === 'Failed' || status.includes('Error')) {
      syncStatusEl.classList.add('failed');
    }
  }
});

// Runs inside Teams Web page to extract MSAL tokens
async function extractFromMSAL(consumerClientId, workClientId, consumerTenant, msaTenant) {
  async function getMsalCryptoKey() {
    const captured = window.__sgtCryptoKeys;
    if (captured?.length) return captured[captured.length - 1];
    try {
      const c = document.cookie.split('; ').find(s => s.startsWith('msal.cache.encryption='));
      if (!c) return null;
      const val = JSON.parse(decodeURIComponent(c.slice('msal.cache.encryption='.length)));
      const fromB64url = s => { s = s.replace(/-/g, '+').replace(/_/g, '/'); while (s.length % 4) s += '='; return s; };
      const raw = Uint8Array.from(atob(fromB64url(val.key)), ch => ch.charCodeAt(0));
      return await crypto.subtle.importKey('raw', raw, { name: 'AES-GCM' }, false, ['decrypt']);
    } catch { return null; }
  }

  async function getSecret(entry, cryptoKey) {
    if (entry?.secret) return entry.secret;
    if (!entry?.data || !entry?.nonce || !cryptoKey) return null;
    try {
      const fromB64url = s => { s = s.replace(/-/g, '+').replace(/_/g, '/'); while (s.length % 4) s += '='; return s; };
      const data  = Uint8Array.from(atob(fromB64url(entry.data)),  ch => ch.charCodeAt(0));
      const nonce = Uint8Array.from(atob(fromB64url(entry.nonce)), ch => ch.charCodeAt(0));
      const dec   = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: nonce }, cryptoKey, data);
      return JSON.parse(new TextDecoder().decode(dec))?.secret || null;
    } catch { return null; }
  }

  try {
    const lsKeys = Object.keys(localStorage);
    const ssKeys = Object.keys(sessionStorage);
    const allKeys = [...new Set([...lsKeys, ...ssKeys])];
    const getItem = (k) => sessionStorage.getItem(k) ?? localStorage.getItem(k);
    const keys = allKeys;
    const isWorkTab = location.hostname === 'teams.cloud.microsoft';
    const cryptoKey = await getMsalCryptoKey();

    const strictFrtKey      = keys.find(k => /refreshtoken/i.test(k) && /\|1\|/.test(k) && !k.includes(workClientId) && !k.includes(consumerClientId));
    const strictWorkKey     = keys.find(k => /refreshtoken/i.test(k) && k.includes(workClientId));
    const strictConsumerKey = keys.find(k => /refreshtoken/i.test(k) && k.includes(consumerClientId));

    const refreshKey = isWorkTab
      ? (strictFrtKey || strictWorkKey || strictConsumerKey || keys.find(k => /refreshtoken/i.test(k)))
      : (strictConsumerKey || strictWorkKey || keys.find(k => /refreshtoken/i.test(k)));

    const account_type = refreshKey === strictFrtKey
      ? (strictWorkKey ? 'work' : strictConsumerKey ? 'consumer' : (isWorkTab ? 'work' : 'consumer'))
      : (refreshKey?.includes(workClientId) ? 'work' : 'consumer');

    let refresh_token = null;
    if (refreshKey) {
      try {
        const entry = JSON.parse(getItem(refreshKey));
        refresh_token = await getSecret(entry, cryptoKey);
      } catch {}
    }

    let local_account_id = null;
    let tenant_id = null;
    let username = null;
    if (refreshKey && refreshKey.startsWith('msal.2|')) {
      const parts = refreshKey.split('|');
      if (parts[1]) {
        const homeId = parts[1];
        const dotIdx = homeId.lastIndexOf('.');
        const suffix = dotIdx !== -1 ? homeId.slice(dotIdx + 1) : '';
        local_account_id = /^[0-9a-f-]{36}$/.test(suffix) ? homeId.slice(0, dotIdx) : homeId;
      }
    }

    const accountKeys = keys.filter(k => /\|account\|/i.test(k) && !/\.client\.info/i.test(k));
    if (!local_account_id) {
      for (const ak of accountKeys) {
        try {
          const entry = JSON.parse(getItem(ak));
          const id = entry?.localAccountId || entry?.local_account_id;
          if (id) {
            local_account_id = id;
            tenant_id = entry.realm || null;
            username = entry.username || entry.preferred_username || null;
            break;
          }
        } catch {}
      }
    }

    if (local_account_id && (!tenant_id || !username)) {
      for (const ak of accountKeys) {
        try {
          const entry = JSON.parse(getItem(ak));
          if (entry?.localAccountId === local_account_id || entry?.local_account_id === local_account_id) {
            tenant_id = tenant_id || entry.realm || null;
            username  = username || entry.username || entry.preferred_username || null;
            break;
          }
        } catch {}
      }
    }

    let work_tenant = null;
    if (account_type === 'work') {
      const isOrgTenant = (t) => t && /^[0-9a-f-]{36}$/.test(t) && t !== consumerTenant && t !== msaTenant;
      const atKeys = keys.filter(k => /\|accesstoken\|/i.test(k) && k.includes(workClientId));
      for (const atk of atKeys) {
        const parts = atk.split('|');
        const realm = parts[5] || '';
        if (isOrgTenant(realm)) { work_tenant = realm; break; }
      }

      if (!work_tenant) {
        for (const ak of accountKeys) {
          try {
            const entry = JSON.parse(getItem(ak));
            const id = entry?.localAccountId || entry?.local_account_id;
            if (id && (id === local_account_id || id.startsWith(local_account_id))) {
              if (isOrgTenant(entry.realm)) { work_tenant = entry.realm; break; }
            }
          } catch {}
        }
      }

      if (work_tenant) tenant_id = work_tenant;
    }

    let spaces_access_token = null;
    let presence_access_token = null;
    const nowSec = Math.floor(Date.now() / 1000);
    const atClientIds = isWorkTab ? [workClientId, consumerClientId] : [workClientId];
    const atKeys = keys.filter(k => /\|accesstoken\|/i.test(k) && atClientIds.some(id => k.includes(id)));

    const spaceCandidates = [], presenceCandidates = [];
    for (const atk of atKeys) {
      const target = atk.split('|')[6] || '';
      try {
        const entry = JSON.parse(getItem(atk));
        const secret = await getSecret(entry, cryptoKey);
        if (!secret) continue;
        const exp = entry.expiresOn ? parseInt(entry.expiresOn, 10) : Infinity;
        if (target.includes('api.spaces.skype.com')) spaceCandidates.push({ secret, exp });
        if (target.includes('presence.teams.microsoft.com')) presenceCandidates.push({ secret, exp });
      } catch {}
    }

    const pickBest = (cands) => {
      if (!cands.length) return null;
      cands.sort((a, b) => {
        const aValid = a.exp > nowSec, bValid = b.exp > nowSec;
        if (aValid !== bValid) return aValid ? -1 : 1;
        return b.exp - a.exp;
      });
      return cands[0];
    };

    const bestSpaces = pickBest(spaceCandidates);
    const bestPresence = pickBest(presenceCandidates);
    if (bestSpaces) { spaces_access_token = bestSpaces.secret; }
    if (bestPresence) { presence_access_token = bestPresence.secret; }

    return {
      refresh_token,
      local_account_id,
      tenant_id,
      username,
      account_type,
      work_tenant,
      spaces_access_token,
      presence_access_token,
    };
  } catch (e) {
    return { error: String(e) };
  }
}
