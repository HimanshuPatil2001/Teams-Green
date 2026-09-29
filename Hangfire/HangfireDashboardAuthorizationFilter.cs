using Hangfire.Annotations;
using Hangfire.Dashboard;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Net;

namespace TeamsGreen.Hangfire
{
    public class HangfireDashboardAuthorizationFilter : IDashboardAuthorizationFilter
    {
        public bool Authorize([NotNull] DashboardContext context)
        {
            // For testing and remote access, allow all requests.
            // WARNING: In a production SaaS env, secure this with auth (e.g. JWT cookies or basic authentication).
            return true;
        }
    }
}
