using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Net.Http;
using TeamsGreen.Data;
using TeamsGreen.Entities;
using TeamsGreen.Services;

namespace TeamsGreen.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class PresenceController : ControllerBase
    {
        private readonly KeepAliveRepository _repository;
        private readonly IMSGraphPresenceService _presenceService;
        private readonly IConfiguration _configuration;

        public PresenceController(
            KeepAliveRepository repository, 
            IMSGraphPresenceService presenceService, 
            IConfiguration configuration)
        {
            _repository = repository;
            _presenceService = presenceService;
            _configuration = configuration;
        }

        /// <summary>
        /// Get public configuration details (e.g., ClientId for OAuth logins).
        /// </summary>
        /// <returns>Public configuration values.</returns>
        [HttpGet("config")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public IActionResult GetConfig()
        {
            var clientId = _configuration["MicrosoftGraph:ClientId"];
            return Ok(new { clientId });
        }

        /// <summary>
        /// Retrieve all registered tenant sessions.
        /// </summary>
        /// <returns>List of tenant sessions.</returns>
        [HttpGet("sessions")]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(List<TenantSession>))]
        public async Task<IActionResult> GetAllSessions()
        {
            var sessions = await _repository.GetAllSessionsAsync();
            return Ok(sessions);
        }

        /// <summary>
        /// Create or update a tenant session with Microsoft Graph authentication details.
        /// </summary>
        /// <param name="request">The tenant session details including tokens and email.</param>
        /// <returns>The created or updated tenant session details.</returns>
        [HttpPost("session")]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(TenantSession))]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> UpsertSession([FromBody] CreateOrUpdateSessionRequest request)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var existingSession = await _repository.GetSessionByEmailAsync(request.UserEmail);

            if (existingSession != null)
            {
                existingSession.TenantId = request.TenantId;
                existingSession.MicrosoftGraphAccessToken = request.MicrosoftGraphAccessToken;
                existingSession.MicrosoftGraphRefreshToken = request.MicrosoftGraphRefreshToken;
                existingSession.TokenExpiryUtc = request.TokenExpiryUtc;
                existingSession.TargetPresenceState = request.TargetPresenceState;
                existingSession.SkypeToken = request.SkypeToken;
                existingSession.IsActive = true; // Auto activate on update

                var updated = await _repository.UpsertSessionAsync(existingSession);
                return Ok(updated);
            }
            else
            {
                var newSession = new TenantSession
                {
                    Id = Guid.NewGuid(),
                    TenantId = request.TenantId,
                    UserEmail = request.UserEmail,
                    MicrosoftGraphAccessToken = request.MicrosoftGraphAccessToken,
                    MicrosoftGraphRefreshToken = request.MicrosoftGraphRefreshToken,
                    TokenExpiryUtc = request.TokenExpiryUtc,
                    TargetPresenceState = request.TargetPresenceState,
                    SkypeToken = request.SkypeToken,
                    IsActive = true
                };

                var created = await _repository.UpsertSessionAsync(newSession);
                return Ok(created);
            }
        }

        /// <summary>
        /// Exchange a Microsoft OAuth2 authorization code for access and refresh tokens,
        /// and activate/upsert the session.
        /// </summary>
        /// <param name="request">The exchange request containing the authorization code and redirect URI.</param>
        /// <returns>The created or updated tenant session.</returns>
        [HttpPost("exchange-code")]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(TenantSession))]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> ExchangeCode([FromBody] ExchangeCodeRequest request)
        {
            if (!ModelState.IsValid || string.IsNullOrEmpty(request.Code) || string.IsNullOrEmpty(request.RedirectUri))
            {
                return BadRequest("Invalid request. 'code' and 'redirectUri' are required.");
            }

            try
            {
                var session = await _presenceService.ExchangeCodeAndCreateSessionAsync(request.Code, request.RedirectUri, request.TargetPresenceState);
                return Ok(session);
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        /// <summary>
        /// Register or update a session directly using an access token extracted by the Chrome Extension.
        /// </summary>
        /// <param name="request">The request containing the extracted access token and target status.</param>
        /// <returns>The registered tenant session.</returns>
        [HttpPost("extension-token")]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(TenantSession))]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> RegisterExtensionToken([FromBody] ExtensionTokenRequest request)
        {
            try
            {
                string? userEmail = request.UserEmail;
                string? tenantId = request.TenantId;
                DateTime? tokenExpiryUtc = request.TokenExpiryUtc;
                string? accessToken = request.AccessToken;
                string? refreshToken = request.RefreshToken;

                // Support both names from payload
                if (string.IsNullOrEmpty(accessToken) && !string.IsNullOrEmpty(request.MicrosoftGraphAccessToken))
                {
                    accessToken = request.MicrosoftGraphAccessToken;
                }
                if (string.IsNullOrEmpty(refreshToken) && !string.IsNullOrEmpty(request.MicrosoftGraphRefreshToken))
                {
                    refreshToken = request.MicrosoftGraphRefreshToken;
                }

                if (!string.IsNullOrEmpty(accessToken))
                {
                    if (string.IsNullOrEmpty(userEmail))
                    {
                        userEmail = GetClaimFromJwt(accessToken, "upn");
                        if (string.IsNullOrEmpty(userEmail))
                        {
                            userEmail = GetClaimFromJwt(accessToken, "unique_name");
                        }
                        if (string.IsNullOrEmpty(userEmail))
                        {
                            userEmail = GetClaimFromJwt(accessToken, "email");
                        }
                    }

                    if (string.IsNullOrEmpty(tenantId))
                    {
                        tenantId = GetClaimFromJwt(accessToken, "tid");
                    }

                    if (tokenExpiryUtc == null)
                    {
                        var expString = GetClaimFromJwt(accessToken, "exp");
                        if (long.TryParse(expString, out var expSeconds))
                        {
                            tokenExpiryUtc = DateTimeOffset.FromUnixTimeSeconds(expSeconds).UtcDateTime;
                        }
                    }
                }

                if (string.IsNullOrEmpty(userEmail))
                {
                    return BadRequest("Failed to resolve user email. Ensure AccessToken is valid or UserEmail is provided.");
                }

                if (string.IsNullOrEmpty(tenantId))
                {
                    tenantId = "common";
                }

                if (tokenExpiryUtc == null)
                {
                    tokenExpiryUtc = DateTime.UtcNow.AddHours(1);
                }

                // Look for existing session by LocalAccountId first, then by UserEmail
                TenantSession? existingSession = null;
                if (!string.IsNullOrEmpty(request.LocalAccountId))
                {
                    existingSession = await _repository.GetSessionByLocalAccountIdAsync(request.LocalAccountId);
                }

                if (existingSession == null)
                {
                    existingSession = await _repository.GetSessionByEmailAsync(userEmail);
                }

                if (existingSession != null)
                {
                    existingSession.TenantId = tenantId;
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        existingSession.MicrosoftGraphAccessToken = accessToken;
                    }
                    if (!string.IsNullOrEmpty(refreshToken))
                    {
                        existingSession.MicrosoftGraphRefreshToken = refreshToken;
                    }
                    existingSession.TokenExpiryUtc = tokenExpiryUtc.Value;
                    existingSession.TargetPresenceState = request.TargetPresenceState;
                    existingSession.IsActive = true;

                    if (!string.IsNullOrEmpty(request.LocalAccountId))
                    {
                        existingSession.LocalAccountId = request.LocalAccountId;
                    }
                    if (!string.IsNullOrEmpty(request.ClientId))
                    {
                        existingSession.ClientId = request.ClientId;
                    }
                    if (!string.IsNullOrEmpty(request.SkypeToken))
                    {
                        existingSession.SkypeToken = request.SkypeToken;
                    }

                    var updated = await _repository.UpsertSessionAsync(existingSession);
                    return Ok(updated);
                }
                else
                {
                    var newSession = new TenantSession
                    {
                        Id = Guid.NewGuid(),
                        TenantId = tenantId,
                        UserEmail = userEmail,
                        MicrosoftGraphAccessToken = accessToken ?? string.Empty,
                        MicrosoftGraphRefreshToken = refreshToken ?? string.Empty,
                        TokenExpiryUtc = tokenExpiryUtc.Value,
                        TargetPresenceState = request.TargetPresenceState,
                        IsActive = true,
                        LocalAccountId = request.LocalAccountId,
                        ClientId = request.ClientId,
                        SkypeToken = request.SkypeToken
                    };

                    var created = await _repository.UpsertSessionAsync(newSession);
                    return Ok(created);
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = $"Failed to register token: {ex.Message}" });
            }
        }

        /// <summary>
        /// Redeem a desktop client auth code server-side to bypass extension CORS constraints.
        /// </summary>
        [HttpPost("redeem-code")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> RedeemCode([FromBody] RedeemCodeRequest request)
        {
            if (string.IsNullOrEmpty(request.Code) || string.IsNullOrEmpty(request.CodeVerifier) || string.IsNullOrEmpty(request.TenantId))
            {
                return BadRequest("Code, codeVerifier, and tenantId are required.");
            }

            try
            {
                var tokenEndpoint = $"https://login.microsoftonline.com/{request.TenantId}/oauth2/v2.0/token";
                var formParams = new Dictionary<string, string>
                {
                    { "grant_type", "authorization_code" },
                    { "client_id", "1fec8e78-bce4-4aaf-ab1b-5451cc387264" }, // Teams Desktop client ID
                    { "code", request.Code },
                    { "redirect_uri", "https://login.microsoftonline.com/common/oauth2/nativeclient" },
                    { "code_verifier", request.CodeVerifier },
                    { "scope", "https://api.spaces.skype.com/user_impersonation offline_access openid profile" }
                };

                using var httpClient = new HttpClient();
                var response = await httpClient.PostAsync(tokenEndpoint, new FormUrlEncodedContent(formParams));
                var responseBody = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    return BadRequest(new { message = $"Token redemption failed: {responseBody}" });
                }

                using var doc = JsonDocument.Parse(responseBody);
                var refreshToken = doc.RootElement.TryGetProperty("refresh_token", out var rtProp) ? rtProp.GetString() : null;
                var accessToken = doc.RootElement.TryGetProperty("access_token", out var atProp) ? atProp.GetString() : null;
                var expiresIn = doc.RootElement.TryGetProperty("expires_in", out var expProp) ? expProp.GetInt32() : 3600;

                return Ok(new
                {
                    refreshToken = refreshToken,
                    accessToken = accessToken,
                    expiresIn = expiresIn
                });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = $"Redemption error: {ex.Message}" });
            }
        }

        /// <summary>
        /// Toggle the active status of a keep-alive session.
        /// </summary>
        /// <param name="id">The Guid identifier of the Tenant Session.</param>
        /// <returns>The updated session indicating active status.</returns>
        [HttpPost("toggle/{id}")]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(TenantSession))]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<IActionResult> ToggleSession(Guid id)
        {
            var session = await _repository.GetSessionByIdAsync(id);
            if (session == null)
            {
                return NotFound(new { message = $"Session with ID {id} not found." });
            }

            var newStatus = await _repository.ToggleSessionAsync(id);
            session.IsActive = newStatus;
            return Ok(session);
        }

        /// <summary>
        /// Manually trigger a keep-alive presence update call for debugging.
        /// </summary>
        /// <param name="id">The Guid identifier of the Tenant Session.</param>
        /// <returns>A message indicating that the ping has been executed.</returns>
        [HttpPost("ping/{id}")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<IActionResult> TriggerPing(Guid id)
        {
            var exists = await _repository.SessionExistsAsync(id);
            if (!exists)
            {
                return NotFound(new { message = $"Session with ID {id} not found." });
            }

            await _presenceService.KeepAliveAsync(id);

            var logs = await _repository.GetLogsBySessionIdAsync(id, 1);
            var latestLog = logs.Count > 0 ? logs[0] : null;

            return Ok(new
            {
                message = "Ping executed.",
                lastResponseCode = latestLog?.HttpStatusCode,
                lastResponseBody = latestLog?.ResponseBody
            });
        }

        /// <summary>
        /// Retrieve the last 50 execution logs for a session to review historical activity.
        /// </summary>
        /// <param name="id">The Guid identifier of the Tenant Session.</param>
        /// <returns>List of latest presence status logs.</returns>
        [HttpGet("logs/{id}")]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(List<PresenceLog>))]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<IActionResult> GetSessionLogs(Guid id)
        {
            var exists = await _repository.SessionExistsAsync(id);
            if (!exists)
            {
                return NotFound(new { message = $"Session with ID {id} not found." });
            }

            var logs = await _repository.GetLogsBySessionIdAsync(id, 50);
            return Ok(logs);
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
                    if (prop.ValueKind == JsonValueKind.Number)
                    {
                        return prop.GetInt64().ToString();
                    }
                    return prop.GetString() ?? string.Empty;
                }
            }
            catch
            {
                // Ignore and return empty
            }
            return string.Empty;
        }
    }

    public class CreateOrUpdateSessionRequest
    {
        public string TenantId { get; set; } = string.Empty;
        public string UserEmail { get; set; } = string.Empty;
        public string MicrosoftGraphAccessToken { get; set; } = string.Empty;
        public string MicrosoftGraphRefreshToken { get; set; } = string.Empty;
        public DateTime TokenExpiryUtc { get; set; }
        public string TargetPresenceState { get; set; } = "Available";
        public string? SkypeToken { get; set; }
    }

    public class ExtensionTokenRequest
    {
        public string? AccessToken { get; set; }
        public string? MicrosoftGraphAccessToken { get; set; }
        public string TargetPresenceState { get; set; } = "Available";
        public string? UserEmail { get; set; }
        public string? TenantId { get; set; }
        public string? RefreshToken { get; set; }
        public string? MicrosoftGraphRefreshToken { get; set; }
        public string? ClientId { get; set; }
        public string? LocalAccountId { get; set; }
        public string? SkypeToken { get; set; }
        public DateTime? TokenExpiryUtc { get; set; }
    }

    public class ExchangeCodeRequest
    {
        public string Code { get; set; } = string.Empty;
        public string RedirectUri { get; set; } = string.Empty;
        public string TargetPresenceState { get; set; } = "Available";
    }

    public class RedeemCodeRequest
    {
        public string Code { get; set; } = string.Empty;
        public string CodeVerifier { get; set; } = string.Empty;
        public string TenantId { get; set; } = string.Empty;
    }
}
