using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using TeamsGreen.Data;
using TeamsGreen.Services;

namespace TeamsGreen.Hangfire
{
    public class KeepAliveBackgroundJob
    {
        private readonly KeepAliveRepository _repository;
        private readonly IMSGraphPresenceService _presenceService;
        private readonly ILogger<KeepAliveBackgroundJob> _logger;

        public KeepAliveBackgroundJob(
            KeepAliveRepository repository,
            IMSGraphPresenceService presenceService,
            ILogger<KeepAliveBackgroundJob> logger)
        {
            _repository = repository;
            _presenceService = presenceService;
            _logger = logger;
        }

        public async Task RunAsync()
        {
            _logger.LogInformation("Hangfire Keep-Alive Job triggered at {Time}", DateTime.UtcNow);

            try
            {
                var activeSessions = await _repository.GetActiveSessionsAsync();

                _logger.LogInformation("Found {Count} active tenant sessions for presence keep-alive.", activeSessions.Count);

                foreach (var session in activeSessions)
                {
                    try
                    {
                        // Process each keep-alive asynchronously in MSGraphPresenceService
                        await _presenceService.KeepAliveAsync(session.Id);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error executing presence keep-alive for session ID {SessionId} (Email: {Email})", session.Id, session.UserEmail);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Critical failure in KeepAliveBackgroundJob root execution.");
            }
        }
    }
}
