using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using TeamsGreen.Data;
using TeamsGreen.Entities;

namespace TeamsGreen.Services
{
    public class MSGraphPresenceService : IMSGraphPresenceService
    {
        private readonly KeepAliveRepository _repository;
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _configuration;
        private readonly ILogger<MSGraphPresenceService> _logger;

        public MSGraphPresenceService(
            KeepAliveRepository repository,
            HttpClient httpClient,
            IConfiguration configuration,
            ILogger<MSGraphPresenceService> logger)
        {
            _repository = repository;
            _httpClient = httpClient;
            _configuration = configuration;
            _logger = logger;
        }

        public async Task KeepAliveAsync(Guid tenantSessionId)
        {
            var session = await _repository.GetSessionByIdAsync(tenantSessionId);
            if (session == null)
            {
                _logger.LogWarning("Session {SessionId} not found.", tenantSessionId);
                return;
            }

            if (!session.IsActive)
            {
                _logger.LogInformation("Session {SessionId} is inactive. Skipping keep-alive.", tenantSessionId);
                return;
            }

            _logger.LogInformation("Starting keep-alive ping for user {UserEmail} ({SessionId}).", session.UserEmail, session.Id);

            string accessToken;
            try
            {
                accessToken = await GetOrRefreshTokenAsync(session);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to refresh access token for user {UserEmail}.", session.UserEmail);
                
                // Record the failure log
                var errorLog = new PresenceLog
                {
                    Id = Guid.NewGuid(),
                    TenantSessionId = session.Id,
                    TimestampUtc = DateTime.UtcNow,
                    HttpStatusCode = (int)HttpStatusCode.Unauthorized,
                    ResponseBody = $"Token refresh failed: {ex.Message}"
                };

                // If authentication is invalid, deactivate the session to prevent continuous failing requests
                if (ex.Message.Contains("invalid_grant", StringComparison.OrdinalIgnoreCase))
                {
                    session.IsActive = false;
                    _logger.LogWarning("Deactivating session {SessionId} due to invalid grant/expired credentials.", session.Id);
                    await _repository.UpsertSessionAsync(session);
                }

                await _repository.AddPresenceLogAsync(errorLog);
                return;
            }

            var aud = GetClaimFromJwt(accessToken, "aud");
            HttpRequestMessage request;

            bool isPersonal = !string.IsNullOrEmpty(session.SkypeToken) || 
                              session.TenantId == "9188040d-6c67-4c5b-b112-36a304b66dad" || 
                              session.TenantId == "f8cdef31-a31e-4b4a-93e4-5f571e91255a";

            if (isPersonal)
            {
                _logger.LogInformation("Using Personal Teams (teams.live.com) forceavailability API for user {UserEmail}.", session.UserEmail);
                var payload = new { availability = session.TargetPresenceState };
                var jsonContent = JsonSerializer.Serialize(payload);
                
                request = new HttpRequestMessage(HttpMethod.Put, "https://teams.live.com/ups/global/v1/me/forceavailability/")
                {
                    Content = new StringContent(jsonContent, Encoding.UTF8, "application/json")
                };
                
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                if (!string.IsNullOrEmpty(session.SkypeToken))
                {
                    request.Headers.TryAddWithoutValidation("x-skypetoken", session.SkypeToken);
                }
                request.Headers.TryAddWithoutValidation("x-ms-client-consumer-type", "teams4life");
                request.Headers.TryAddWithoutValidation("x-ms-client-caller", "TeamsClient:setPresenceAvailability");
                request.Headers.TryAddWithoutValidation("x-ms-client-user-agent", "Teams-V2-Web");
                request.Headers.TryAddWithoutValidation("x-ms-client-type", "cdlworker");
                request.Headers.TryAddWithoutValidation("x-ms-client-version", "1415/26080200346");
                request.Headers.TryAddWithoutValidation("x-ms-endpoint-id", session.Id.ToString());
                request.Headers.TryAddWithoutValidation("x-ms-session-id", session.Id.ToString());
                request.Headers.TryAddWithoutValidation("behavioroverride", "redirectAs404");
                request.Headers.TryAddWithoutValidation("Referer", "https://teams.live.com/");
            }
            else if (aud.Contains("presence.teams.microsoft.com", StringComparison.OrdinalIgnoreCase) || 
                     aud.Contains("skype.com", StringComparison.OrdinalIgnoreCase) || 
                     aud.Contains("skype", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogInformation("Using Enterprise Teams forceavailability API for user {UserEmail}.", session.UserEmail);
                var payload = new { availability = session.TargetPresenceState };
                var jsonContent = JsonSerializer.Serialize(payload);

                request = new HttpRequestMessage(HttpMethod.Put, "https://presence.teams.microsoft.com/v1/me/forceavailability/")
                {
                    Content = new StringContent(jsonContent, Encoding.UTF8, "application/json")
                };

                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                if (!string.IsNullOrEmpty(session.SkypeToken))
                {
                    request.Headers.TryAddWithoutValidation("x-skypetoken", session.SkypeToken);
                }
                request.Headers.TryAddWithoutValidation("x-ms-client-user-agent", "Teams-V2-Web");
                request.Headers.TryAddWithoutValidation("x-ms-client-type", "cdlworker");
                request.Headers.TryAddWithoutValidation("x-ms-client-caller", "TeamsClient:setPresenceAvailability");
                request.Headers.TryAddWithoutValidation("x-ms-client-version", "1415/26080200616");
                request.Headers.TryAddWithoutValidation("x-ms-endpoint-id", session.Id.ToString());
                request.Headers.TryAddWithoutValidation("x-ms-session-id", session.Id.ToString());
                request.Headers.TryAddWithoutValidation("behavioroverride", "redirectAs404");
                request.Headers.TryAddWithoutValidation("Referer", "https://teams.cloud.microsoft/");
                request.Headers.TryAddWithoutValidation("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0");
            }
            else
            {
                _logger.LogInformation("Using Microsoft Graph presence API for user {UserEmail}.", session.UserEmail);
                var presencePayload = new SetPresenceRequest
                {
                    SessionId = session.Id.ToString(),
                    Availability = session.TargetPresenceState,
                    Activity = session.TargetPresenceState,
                    ExpirationDuration = "PT1H" // 1 hour duration
                };
                var jsonContent = JsonSerializer.Serialize(presencePayload);
                request = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/me/presence/setPresence")
                {
                    Content = new StringContent(jsonContent, Encoding.UTF8, "application/json")
                };
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            }

            HttpResponseMessage response;
            string responseBody = string.Empty;
            try
            {
                response = await _httpClient.SendAsync(request);
                responseBody = await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "HTTP exception during keep-alive ping for user {UserEmail}.", session.UserEmail);
                
                var networkErrorLog = new PresenceLog
                {
                    Id = Guid.NewGuid(),
                    TenantSessionId = session.Id,
                    TimestampUtc = DateTime.UtcNow,
                    HttpStatusCode = (int)HttpStatusCode.InternalServerError,
                    ResponseBody = $"Network error: {ex.Message}"
                };
                
                await _repository.AddPresenceLogAsync(networkErrorLog);
                return;
            }

            _logger.LogInformation("Keep-alive response for user {UserEmail}: HttpStatusCode={StatusCode}", session.UserEmail, response.StatusCode);

            // Save presence log
            var log = new PresenceLog
            {
                Id = Guid.NewGuid(),
                TenantSessionId = session.Id,
                TimestampUtc = DateTime.UtcNow,
                HttpStatusCode = (int)response.StatusCode,
                ResponseBody = string.IsNullOrWhiteSpace(responseBody) ? $"HTTP Status {response.StatusCode}" : responseBody
            };

            await _repository.AddPresenceLogAsync(log);
        }

        private async Task<string> GetOrRefreshTokenAsync(TenantSession session)
        {
            // If the token is still valid (with a 5-minute buffer) and not empty, return it directly
            if (!string.IsNullOrEmpty(session.MicrosoftGraphAccessToken) && session.TokenExpiryUtc > DateTime.UtcNow.AddMinutes(5))
            {
                return session.MicrosoftGraphAccessToken;
            }

            // If we don't have a refresh token (Extension flow), we cannot perform a silent refresh.
            if (string.IsNullOrEmpty(session.MicrosoftGraphRefreshToken))
            {
                _logger.LogWarning("Access token expired and no refresh token is available for session {SessionId}. Deactivating.", session.Id);
                session.IsActive = false;
                await _repository.UpsertSessionAsync(session);
                throw new Exception("Access token expired. Please open Teams Web in your browser to automatically renew your session.");
            }

            _logger.LogInformation("Access token expired or expiring soon for user {UserEmail}. Refreshing...", session.UserEmail);

            var tenantId = session.TenantId;
            var clientId = session.ClientId;
            bool isPublicClient = false;
            string scope = "https://graph.microsoft.com/Presence.ReadWrite";

            if (string.IsNullOrEmpty(clientId))
            {
                clientId = _configuration["MicrosoftGraph:ClientId"] ?? throw new InvalidOperationException("MicrosoftGraph:ClientId is not configured.");
            }

            // Check if it's one of the first-party Microsoft public clients
            if (clientId == "1fec8e78-bce4-4aaf-ab1b-5451cc387264" || // Desktop Client
                clientId == "5e3ce6c0-2b1f-4285-8d4b-75ee78787346" || // Work Web Client
                clientId == "4b3e8f46-56d3-427f-b1e2-d239b2ea6bca")   // Consumer Web Client
            {
                isPublicClient = true;
                if (clientId == "1fec8e78-bce4-4aaf-ab1b-5451cc387264")
                {
                    scope = "https://api.spaces.skype.com/user_impersonation offline_access openid profile";
                }
            }

            var tokenEndpoint = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            var formParams = new Dictionary<string, string>
            {
                { "grant_type", "refresh_token" },
                { "client_id", clientId },
                { "refresh_token", session.MicrosoftGraphRefreshToken },
                { "scope", scope }
            };

            if (!isPublicClient)
            {
                var clientSecret = _configuration["MicrosoftGraph:ClientSecret"];
                if (!string.IsNullOrEmpty(clientSecret))
                {
                    formParams.Add("client_secret", clientSecret);
                }
            }

            var requestContent = new FormUrlEncodedContent(formParams);
            var httpRequest = new HttpRequestMessage(HttpMethod.Post, tokenEndpoint)
            {
                Content = requestContent
            };

            if (!string.IsNullOrEmpty(session.LocalAccountId))
            {
                httpRequest.Headers.Add("X-AnchorMailbox", $"Oid:{session.LocalAccountId}@{tenantId}");
            }

            var response = await _httpClient.SendAsync(httpRequest);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("Token refresh failed. StatusCode={StatusCode}, Body={Body}", response.StatusCode, responseBody);
                throw new Exception($"Token refresh API returned status code {response.StatusCode}. Response: {responseBody}");
            }

            var tokenResponse = JsonSerializer.Deserialize<OAuthTokenResponse>(responseBody) 
                                ?? throw new Exception("Failed to deserialize OAuth token response.");

            // Update session details
            session.MicrosoftGraphAccessToken = tokenResponse.AccessToken;
            if (!string.IsNullOrEmpty(tokenResponse.RefreshToken))
            {
                session.MicrosoftGraphRefreshToken = tokenResponse.RefreshToken;
            }
            session.TokenExpiryUtc = DateTime.UtcNow.AddSeconds(tokenResponse.ExpiresIn);

            await _repository.UpsertSessionAsync(session);

            _logger.LogInformation("Successfully refreshed access token for user {UserEmail}.", session.UserEmail);
            return session.MicrosoftGraphAccessToken;
        }

        public async Task<TenantSession> ExchangeCodeAndCreateSessionAsync(string code, string redirectUri, string targetPresenceState)
        {
            _logger.LogInformation("Exchanging authorization code for tokens.");

            var clientId = _configuration["MicrosoftGraph:ClientId"] ?? throw new InvalidOperationException("MicrosoftGraph:ClientId is not configured.");
            var clientSecret = _configuration["MicrosoftGraph:ClientSecret"] ?? throw new InvalidOperationException("MicrosoftGraph:ClientSecret is not configured.");

            var tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

            var formParams = new Dictionary<string, string>
            {
                { "grant_type", "authorization_code" },
                { "client_id", clientId },
                { "client_secret", clientSecret },
                { "code", code },
                { "redirect_uri", redirectUri },
                { "scope", "offline_access https://graph.microsoft.com/.default" }
            };

            var requestContent = new FormUrlEncodedContent(formParams);
            var response = await _httpClient.PostAsync(tokenEndpoint, requestContent);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("Authorization code exchange failed. StatusCode={StatusCode}, Body={Body}", response.StatusCode, responseBody);
                throw new Exception($"Failed to exchange authorization code: {responseBody}");
            }

            var tokenResponse = JsonSerializer.Deserialize<OAuthTokenResponse>(responseBody) 
                                ?? throw new Exception("Failed to deserialize OAuth token response.");

            // Retrieve user profile details from Microsoft Graph
            var meRequest = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me");
            meRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", tokenResponse.AccessToken);

            var meResponse = await _httpClient.SendAsync(meRequest);
            var meResponseBody = await meResponse.Content.ReadAsStringAsync();

            if (!meResponse.IsSuccessStatusCode)
            {
                _logger.LogError("Failed to retrieve user profile from Microsoft Graph. StatusCode={StatusCode}, Body={Body}", meResponse.StatusCode, meResponseBody);
                throw new Exception($"Failed to retrieve user profile: {meResponseBody}");
            }

            using var meJson = JsonDocument.Parse(meResponseBody);
            var userEmail = string.Empty;
            if (meJson.RootElement.TryGetProperty("mail", out var mailProp) && !string.IsNullOrEmpty(mailProp.GetString()))
            {
                userEmail = mailProp.GetString()!;
            }
            else if (meJson.RootElement.TryGetProperty("userPrincipalName", out var upnProp))
            {
                userEmail = upnProp.GetString()!;
            }

            if (string.IsNullOrEmpty(userEmail))
            {
                throw new Exception("Could not retrieve email or user principal name from Graph user profile.");
            }

            // Extract the Tenant ID from the Access Token (tid claim)
            var tenantId = GetClaimFromJwt(tokenResponse.AccessToken, "tid");
            if (string.IsNullOrEmpty(tenantId))
            {
                tenantId = "common";
            }

            // Look for existing session
            var existingSession = await _repository.GetSessionByEmailAsync(userEmail);

            if (existingSession != null)
            {
                existingSession.TenantId = tenantId;
                existingSession.MicrosoftGraphAccessToken = tokenResponse.AccessToken;
                existingSession.MicrosoftGraphRefreshToken = tokenResponse.RefreshToken;
                existingSession.TokenExpiryUtc = DateTime.UtcNow.AddSeconds(tokenResponse.ExpiresIn);
                existingSession.TargetPresenceState = targetPresenceState;
                existingSession.IsActive = true;

                return await _repository.UpsertSessionAsync(existingSession);
            }
            else
            {
                var newSession = new TenantSession
                {
                    Id = Guid.NewGuid(),
                    TenantId = tenantId,
                    UserEmail = userEmail,
                    MicrosoftGraphAccessToken = tokenResponse.AccessToken,
                    MicrosoftGraphRefreshToken = tokenResponse.RefreshToken,
                    TokenExpiryUtc = DateTime.UtcNow.AddSeconds(tokenResponse.ExpiresIn),
                    TargetPresenceState = targetPresenceState,
                    IsActive = true
                };

                return await _repository.UpsertSessionAsync(newSession);
            }
        }

        private static string GetClaimFromJwt(string jwt, string claimType)
        {
            try
            {
                var parts = jwt.Split('.');
                if (parts.Length < 2) return string.Empty;
                var payload = parts[1];
                
                payload = payload.Replace('-', '+').Replace('_', '/');
                switch (payload.Length % 4)
                {
                    case 2: payload += "=="; break;
                    case 3: payload += "="; break;
                }
                
                var bytes = Convert.FromBase64String(payload);
                var json = System.Text.Encoding.UTF8.GetString(bytes);
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty(claimType, out var prop))
                {
                    return prop.GetString() ?? string.Empty;
                }
            }
            catch
            {
                // Ignore and return empty
            }
            return string.Empty;
        }

        private class SetPresenceRequest
        {
            [JsonPropertyName("sessionId")]
            public string SessionId { get; set; } = string.Empty;

            [JsonPropertyName("availability")]
            public string Availability { get; set; } = string.Empty;

            [JsonPropertyName("activity")]
            public string Activity { get; set; } = string.Empty;

            [JsonPropertyName("expirationDuration")]
            public string ExpirationDuration { get; set; } = string.Empty;
        }

        private class OAuthTokenResponse
        {
            [JsonPropertyName("access_token")]
            public string AccessToken { get; set; } = string.Empty;

            [JsonPropertyName("refresh_token")]
            public string RefreshToken { get; set; } = string.Empty;

            [JsonPropertyName("expires_in")]
            public int ExpiresIn { get; set; }
        }
    }
}
