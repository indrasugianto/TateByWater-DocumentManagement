using System.Data;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Options;
using TBCMSDropboxBridge.Models;

namespace TBCMSDropboxBridge.Services;

public interface IDropboxTokenService
{
    // Returns a non-expired access token, refreshing first if within the skew
    // window. Throws InvalidOperationException("Bridge not configured...") if
    // the token table is empty.
    Task<string?> GetValidAccessTokenAsync(CancellationToken ct = default);

    // UPSERTs the singleton row (TokenID = 1). AccessToken/RefreshToken are
    // encrypted at rest with machine-scope Data Protection; AccountEmail is
    // stored in plaintext.
    Task SaveTokensAsync(string accessToken, string refreshToken,
                         int expiresInSeconds, string accountEmail,
                         string setupByUser, CancellationToken ct = default);

    Task<bool> HasTokenAsync(CancellationToken ct = default);

    // Returns the email STORED in tblDropboxServiceToken — no network call.
    // Used by RefreshAsync to preserve the email across a refresh.
    Task<string?> GetAccountEmailAsync(CancellationToken ct = default);

    // Makes a live users/get_current_account call to prove the token is still
    // valid (not revoked). Used by /api/status. Throws on auth failure.
    Task<string?> VerifyLiveAccountEmailAsync(CancellationToken ct = default);
}

public sealed class DropboxTokenService : IDropboxTokenService
{
    private const string TokenUrl   = "https://api.dropboxapi.com/oauth2/token";
    private const string AccountUrl = "https://api.dropboxapi.com/2/users/get_current_account";
    // Refresh slightly before actual expiry so in-flight calls don't get a token
    // that expires mid-request.
    private static readonly TimeSpan ExpirySkew = TimeSpan.FromMinutes(5);

    // MUST be static: DropboxTokenService is registered scoped, so a per-instance
    // semaphore would give every request its own lock and serialize nothing (G39).
    // One process-wide lock around refresh+save prevents concurrent near-expiry
    // requests from double-refreshing / racing the SQL write.
    private static readonly SemaphoreSlim RefreshLock = new(1, 1);

    private readonly SqlConnection _cn;
    private readonly IDataProtector _protector;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly DropboxOptions _cfg;
    private readonly ILogger<DropboxTokenService> _log;

    public DropboxTokenService(
        SqlConnection cn,
        IDataProtectionProvider dpProvider,
        IHttpClientFactory httpClientFactory,
        IOptions<DropboxOptions> cfg,
        ILogger<DropboxTokenService> log)
    {
        _cn = cn;
        _protector = dpProvider.CreateProtector("TBCMSDropboxBridge.ServiceToken");
        _httpClientFactory = httpClientFactory;
        _cfg = cfg.Value;
        _log = log;
    }

    private sealed record TokenRow(
        string AccessCipher, string RefreshCipher, DateTime ExpiresAtUtc, string? AccountEmail);

    // ----- public API --------------------------------------------------------

    public async Task<bool> HasTokenAsync(CancellationToken ct = default)
        => await LoadRowAsync(ct) is not null;

    public async Task<string?> GetAccountEmailAsync(CancellationToken ct = default)
        => (await LoadRowAsync(ct))?.AccountEmail;

    public async Task<string?> GetValidAccessTokenAsync(CancellationToken ct = default)
    {
        var row = await LoadRowAsync(ct)
                  ?? throw new InvalidOperationException("Bridge not configured — run setup first");

        if (!IsExpiring(row))
            return _protector.Unprotect(row.AccessCipher);

        // Near expiry: serialize the refresh across all concurrent requests.
        await RefreshLock.WaitAsync(ct);
        try
        {
            // Re-load inside the lock: another request may have just refreshed.
            row = await LoadRowAsync(ct)
                  ?? throw new InvalidOperationException("Bridge not configured — run setup first");
            if (IsExpiring(row))
            {
                await RefreshAsync(_protector.Unprotect(row.RefreshCipher), ct);
                row = await LoadRowAsync(ct)
                      ?? throw new InvalidOperationException("Bridge not configured — run setup first");
            }
        }
        finally
        {
            RefreshLock.Release();
        }

        return _protector.Unprotect(row.AccessCipher);
    }

    public async Task<string?> VerifyLiveAccountEmailAsync(CancellationToken ct = default)
    {
        var token = await GetValidAccessTokenAsync(ct);
        using var http = _httpClientFactory.CreateClient();
        using var req = new HttpRequestMessage(HttpMethod.Post, AccountUrl);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        // users/get_current_account takes a null body but rejects a JSON content
        // type; send no content.
        using var resp = await http.SendAsync(req, ct);
        if (resp.StatusCode == HttpStatusCode.Unauthorized)
            throw new HttpRequestException("Dropbox token rejected (401) — token may be revoked.",
                                           null, HttpStatusCode.Unauthorized);
        resp.EnsureSuccessStatusCode();
        var acct = await resp.Content.ReadFromJsonAsync<JsonElement>(ct);
        return acct.TryGetProperty("email", out var e) ? e.GetString() : null;
    }

    public async Task SaveTokensAsync(string accessToken, string refreshToken,
                                      int expiresInSeconds, string accountEmail,
                                      string setupByUser, CancellationToken ct = default)
    {
        var accessCipher  = _protector.Protect(accessToken);
        var refreshCipher = _protector.Protect(refreshToken);
        var expiresAtUtc  = DateTime.UtcNow.AddSeconds(expiresInSeconds);

        // UPSERT the singleton row. UPDATE first; INSERT only if no row existed.
        const string sql = @"
UPDATE dbo.tblDropboxServiceToken
   SET AccessToken = @access, RefreshToken = @refresh, ExpiresAtUtc = @expires,
       AccountEmail = @email, UpdatedAtUtc = SYSUTCDATETIME(), SetupByUser = @user
 WHERE TokenID = 1;
IF @@ROWCOUNT = 0
   INSERT INTO dbo.tblDropboxServiceToken
       (TokenID, AccessToken, RefreshToken, ExpiresAtUtc, AccountEmail, UpdatedAtUtc, SetupByUser)
   VALUES (1, @access, @refresh, @expires, @email, SYSUTCDATETIME(), @user);";

        if (_cn.State != ConnectionState.Open) await _cn.OpenAsync(ct);
        try
        {
            using var cmd = new SqlCommand(sql, _cn);
            cmd.Parameters.Add("@access",  SqlDbType.NVarChar, -1).Value = accessCipher;
            cmd.Parameters.Add("@refresh", SqlDbType.NVarChar, -1).Value = refreshCipher;
            cmd.Parameters.Add("@expires", SqlDbType.DateTime2).Value    = expiresAtUtc;
            cmd.Parameters.Add("@email",   SqlDbType.NVarChar, 200).Value =
                string.IsNullOrEmpty(accountEmail) ? DBNull.Value : accountEmail;
            cmd.Parameters.Add("@user",    SqlDbType.NVarChar, 200).Value =
                string.IsNullOrEmpty(setupByUser) ? DBNull.Value : setupByUser;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        finally
        {
            await _cn.CloseAsync();
        }
    }

    // ----- internals ---------------------------------------------------------

    private static bool IsExpiring(TokenRow row) => row.ExpiresAtUtc - ExpirySkew <= DateTime.UtcNow;

    private async Task<TokenRow?> LoadRowAsync(CancellationToken ct)
    {
        const string sql = @"
SELECT TOP 1 AccessToken, RefreshToken, ExpiresAtUtc, AccountEmail
FROM dbo.tblDropboxServiceToken
ORDER BY UpdatedAtUtc DESC;";

        if (_cn.State != ConnectionState.Open) await _cn.OpenAsync(ct);
        try
        {
            using var cmd = new SqlCommand(sql, _cn);
            using var rdr = await cmd.ExecuteReaderAsync(ct);
            if (!await rdr.ReadAsync(ct)) return null;
            return new TokenRow(
                rdr.GetString(0),
                rdr.GetString(1),
                rdr.GetDateTime(2),
                rdr.IsDBNull(3) ? null : rdr.GetString(3));
        }
        finally
        {
            await _cn.CloseAsync();
        }
    }

    private async Task RefreshAsync(string refreshToken, CancellationToken ct)
    {
        using var http = _httpClientFactory.CreateClient();
        var form = new Dictionary<string, string>
        {
            ["grant_type"]    = "refresh_token",
            ["refresh_token"] = refreshToken,
            ["client_id"]     = _cfg.AppKey,
            ["client_secret"] = _cfg.AppSecret,
        };
        using var resp = await http.PostAsync(TokenUrl, new FormUrlEncodedContent(form), ct);
        resp.EnsureSuccessStatusCode();
        var json = await resp.Content.ReadFromJsonAsync<JsonElement>(ct);
        var newAccess    = json.GetProperty("access_token").GetString()!;
        var newExpiresIn = json.GetProperty("expires_in").GetInt32();

        // Dropbox does not rotate the refresh token for offline_access grants —
        // keep the existing one. AND preserve the stored account email: a live
        // users/get_current_account call is impossible here (we refresh precisely
        // because the token is expired), and passing "" would wipe the email on
        // every refresh.
        var existingEmail = await GetAccountEmailAsync(ct) ?? "";
        await SaveTokensAsync(newAccess, refreshToken, newExpiresIn,
                              existingEmail, "auto-refresh", ct);
        _log.LogInformation("Service token refreshed; expires in {Seconds}s.", newExpiresIn);
    }
}
