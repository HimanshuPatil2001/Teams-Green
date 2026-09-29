using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using TeamsGreen.Entities;

namespace TeamsGreen.Data
{
    public class KeepAliveRepository
    {
        private readonly string _connectionString;

        public KeepAliveRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public async Task<List<TenantSession>> GetActiveSessionsAsync()
        {
            var list = new List<TenantSession>();
            const string query = "SELECT Id, TenantId, UserEmail, MicrosoftGraphAccessToken, MicrosoftGraphRefreshToken, TokenExpiryUtc, IsActive, TargetPresenceState, LocalAccountId, ClientId, SkypeToken FROM TenantSessions WHERE IsActive = 1";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                await connection.OpenAsync();
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        list.Add(MapSession(reader));
                    }
                }
            }
            return list;
        }

        public async Task<List<TenantSession>> GetAllSessionsAsync()
        {
            var list = new List<TenantSession>();
            const string query = "SELECT Id, TenantId, UserEmail, MicrosoftGraphAccessToken, MicrosoftGraphRefreshToken, TokenExpiryUtc, IsActive, TargetPresenceState, LocalAccountId, ClientId, SkypeToken FROM TenantSessions";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                await connection.OpenAsync();
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        list.Add(MapSession(reader));
                    }
                }
            }
            return list;
        }

        public async Task<TenantSession?> GetSessionByIdAsync(Guid id)
        {
            const string query = "SELECT Id, TenantId, UserEmail, MicrosoftGraphAccessToken, MicrosoftGraphRefreshToken, TokenExpiryUtc, IsActive, TargetPresenceState, LocalAccountId, ClientId, SkypeToken FROM TenantSessions WHERE Id = @Id";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Id", id);
                await connection.OpenAsync();
                using (var reader = await command.ExecuteReaderAsync())
                {
                    if (await reader.ReadAsync())
                    {
                        return MapSession(reader);
                    }
                }
            }
            return null;
        }

        public async Task<TenantSession?> GetSessionByEmailAsync(string email)
        {
            const string query = "SELECT Id, TenantId, UserEmail, MicrosoftGraphAccessToken, MicrosoftGraphRefreshToken, TokenExpiryUtc, IsActive, TargetPresenceState, LocalAccountId, ClientId, SkypeToken FROM TenantSessions WHERE LOWER(UserEmail) = LOWER(@Email)";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Email", email);
                await connection.OpenAsync();
                using (var reader = await command.ExecuteReaderAsync())
                {
                    if (await reader.ReadAsync())
                    {
                        return MapSession(reader);
                    }
                }
            }
            return null;
        }

        public async Task<TenantSession?> GetSessionByLocalAccountIdAsync(string localAccountId)
        {
            const string query = "SELECT Id, TenantId, UserEmail, MicrosoftGraphAccessToken, MicrosoftGraphRefreshToken, TokenExpiryUtc, IsActive, TargetPresenceState, LocalAccountId, ClientId, SkypeToken FROM TenantSessions WHERE LocalAccountId = @LocalAccountId";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@LocalAccountId", localAccountId);
                await connection.OpenAsync();
                using (var reader = await command.ExecuteReaderAsync())
                {
                    if (await reader.ReadAsync())
                    {
                        return MapSession(reader);
                    }
                }
            }
            return null;
        }

        public async Task<TenantSession> UpsertSessionAsync(TenantSession session)
        {
            const string query = @"
                MERGE TenantSessions AS target
                USING (SELECT @Id AS Id, @TenantId AS TenantId, @UserEmail AS UserEmail, 
                              @MicrosoftGraphAccessToken AS MicrosoftGraphAccessToken, 
                              @MicrosoftGraphRefreshToken AS MicrosoftGraphRefreshToken, 
                              @TokenExpiryUtc AS TokenExpiryUtc, @IsActive AS IsActive, 
                              @TargetPresenceState AS TargetPresenceState,
                              @LocalAccountId AS LocalAccountId, @ClientId AS ClientId,
                              @SkypeToken AS SkypeToken) AS source
                ON (LOWER(target.UserEmail) = LOWER(source.UserEmail))
                WHEN MATCHED THEN
                    UPDATE SET TenantId = source.TenantId, 
                               MicrosoftGraphAccessToken = source.MicrosoftGraphAccessToken, 
                               MicrosoftGraphRefreshToken = source.MicrosoftGraphRefreshToken, 
                               TokenExpiryUtc = source.TokenExpiryUtc, 
                               TargetPresenceState = source.TargetPresenceState, 
                               IsActive = source.IsActive,
                               LocalAccountId = source.LocalAccountId,
                               ClientId = source.ClientId,
                               SkypeToken = source.SkypeToken
                WHEN NOT MATCHED THEN
                    INSERT (Id, TenantId, UserEmail, MicrosoftGraphAccessToken, MicrosoftGraphRefreshToken, TokenExpiryUtc, IsActive, TargetPresenceState, LocalAccountId, ClientId, SkypeToken)
                    VALUES (source.Id, source.TenantId, source.UserEmail, source.MicrosoftGraphAccessToken, source.MicrosoftGraphRefreshToken, source.TokenExpiryUtc, source.IsActive, source.TargetPresenceState, source.LocalAccountId, source.ClientId, source.SkypeToken);";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Id", session.Id);
                command.Parameters.AddWithValue("@TenantId", session.TenantId);
                command.Parameters.AddWithValue("@UserEmail", session.UserEmail);
                command.Parameters.AddWithValue("@MicrosoftGraphAccessToken", session.MicrosoftGraphAccessToken);
                command.Parameters.AddWithValue("@MicrosoftGraphRefreshToken", session.MicrosoftGraphRefreshToken);
                command.Parameters.AddWithValue("@TokenExpiryUtc", session.TokenExpiryUtc);
                command.Parameters.AddWithValue("@IsActive", session.IsActive);
                command.Parameters.AddWithValue("@TargetPresenceState", session.TargetPresenceState);
                command.Parameters.AddWithValue("@LocalAccountId", (object?)session.LocalAccountId ?? DBNull.Value);
                command.Parameters.AddWithValue("@ClientId", (object?)session.ClientId ?? DBNull.Value);
                command.Parameters.AddWithValue("@SkypeToken", (object?)session.SkypeToken ?? DBNull.Value);

                await connection.OpenAsync();
                await command.ExecuteNonQueryAsync();
            }

            // Return the updated session from the DB
            return await GetSessionByEmailAsync(session.UserEmail) ?? session;
        }

        public async Task<bool> ToggleSessionAsync(Guid id)
        {
            const string query = @"
                UPDATE TenantSessions SET IsActive = ~IsActive WHERE Id = @Id;
                SELECT IsActive FROM TenantSessions WHERE Id = @Id;";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Id", id);
                await connection.OpenAsync();
                var result = await command.ExecuteScalarAsync();
                return result != null && (bool)result;
            }
        }

        public async Task AddPresenceLogAsync(PresenceLog log)
        {
            const string query = "INSERT INTO PresenceLogs (Id, TenantSessionId, TimestampUtc, HttpStatusCode, ResponseBody) VALUES (@Id, @TenantSessionId, @TimestampUtc, @HttpStatusCode, @ResponseBody)";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Id", log.Id);
                command.Parameters.AddWithValue("@TenantSessionId", log.TenantSessionId);
                command.Parameters.AddWithValue("@TimestampUtc", log.TimestampUtc);
                command.Parameters.AddWithValue("@HttpStatusCode", log.HttpStatusCode);
                command.Parameters.AddWithValue("@ResponseBody", log.ResponseBody);

                await connection.OpenAsync();
                await command.ExecuteNonQueryAsync();
            }
        }

        public async Task<List<PresenceLog>> GetLogsBySessionIdAsync(Guid sessionId, int limit)
        {
            var list = new List<PresenceLog>();
            var query = $"SELECT TOP (@Limit) Id, TenantSessionId, TimestampUtc, HttpStatusCode, ResponseBody FROM PresenceLogs WHERE TenantSessionId = @SessionId ORDER BY TimestampUtc DESC";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@SessionId", sessionId);
                command.Parameters.AddWithValue("@Limit", limit);

                await connection.OpenAsync();
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        list.Add(new PresenceLog
                        {
                            Id = reader.GetGuid(reader.GetOrdinal("Id")),
                            TenantSessionId = reader.GetGuid(reader.GetOrdinal("TenantSessionId")),
                            TimestampUtc = reader.GetDateTime(reader.GetOrdinal("TimestampUtc")),
                            HttpStatusCode = reader.GetInt32(reader.GetOrdinal("HttpStatusCode")),
                            ResponseBody = reader.GetString(reader.GetOrdinal("ResponseBody"))
                        });
                    }
                }
            }
            return list;
        }

        public async Task<bool> SessionExistsAsync(Guid id)
        {
            const string query = "SELECT COUNT(1) FROM TenantSessions WHERE Id = @Id";

            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Id", id);
                await connection.OpenAsync();
                var count = (int)await command.ExecuteScalarAsync()!;
                return count > 0;
            }
        }

        private static TenantSession MapSession(SqlDataReader reader)
        {
            var localAccountIdOrdinal = reader.GetOrdinal("LocalAccountId");
            var clientIdOrdinal = reader.GetOrdinal("ClientId");
            int skypeTokenOrdinal = -1;
            try { skypeTokenOrdinal = reader.GetOrdinal("SkypeToken"); } catch { }

            return new TenantSession
            {
                Id = reader.GetGuid(reader.GetOrdinal("Id")),
                TenantId = reader.GetString(reader.GetOrdinal("TenantId")),
                UserEmail = reader.GetString(reader.GetOrdinal("UserEmail")),
                MicrosoftGraphAccessToken = reader.GetString(reader.GetOrdinal("MicrosoftGraphAccessToken")),
                MicrosoftGraphRefreshToken = reader.GetString(reader.GetOrdinal("MicrosoftGraphRefreshToken")),
                TokenExpiryUtc = reader.GetDateTime(reader.GetOrdinal("TokenExpiryUtc")),
                IsActive = reader.GetBoolean(reader.GetOrdinal("IsActive")),
                TargetPresenceState = reader.GetString(reader.GetOrdinal("TargetPresenceState")),
                LocalAccountId = reader.IsDBNull(localAccountIdOrdinal) ? null : reader.GetString(localAccountIdOrdinal),
                ClientId = reader.IsDBNull(clientIdOrdinal) ? null : reader.GetString(clientIdOrdinal),
                SkypeToken = (skypeTokenOrdinal >= 0 && !reader.IsDBNull(skypeTokenOrdinal)) ? reader.GetString(skypeTokenOrdinal) : null
            };
        }
    }
}
