using System;

namespace TeamsGreen.Entities
{
    public class PresenceLog
    {
        public Guid Id { get; set; }
        public Guid TenantSessionId { get; set; }
        public DateTime TimestampUtc { get; set; }
        public int HttpStatusCode { get; set; }
        public string ResponseBody { get; set; } = string.Empty;

        // Navigation property
        public TenantSession? TenantSession { get; set; }
    }
}
