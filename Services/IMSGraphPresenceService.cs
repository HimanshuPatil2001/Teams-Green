using System;
using System.Threading.Tasks;
using TeamsGreen.Entities;

namespace TeamsGreen.Services
{
    public interface IMSGraphPresenceService
    {
        Task KeepAliveAsync(Guid tenantSessionId);
        Task<TenantSession> ExchangeCodeAndCreateSessionAsync(string code, string redirectUri, string targetPresenceState);
    }
}
