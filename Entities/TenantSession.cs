using System;

namespace TeamsGreen.Entities
{
    public class TenantSession
    {
        public Guid Id { get; set; }
        public string TenantId { get; set; } = string.Empty;
        public string UserEmail { get; set; } = string.Empty;
        public string MicrosoftGraphAccessToken { get; set; } = string.Empty;
        public string MicrosoftGraphRefreshToken { get; set; } = string.Empty;
        public DateTime TokenExpiryUtc { get; set; }
        public bool IsActive { get; set; }
        public string TargetPresenceState { get; set; } = "Available";
        public string? LocalAccountId { get; set; }
        public string? ClientId { get; set; }
        public string? SkypeToken { get; set; }
    }
}
