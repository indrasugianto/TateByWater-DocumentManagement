using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Options;
using TBCMSDropboxBridge.Models;

namespace TBCMSDropboxBridge.Services;

// Thin wrapper over the Dropbox API. Every call:
//   1. injects the Dropbox-API-Path-Root namespace header (team tree),
//   2. gets a valid access token from IDropboxTokenService (refresh is handled
//      there; callers never see raw tokens),
//   3. translates Dropbox errors into DropboxApiException so the global handler
//      maps them to the HTTP statuses VBA expects (B.7).
public sealed class DropboxApiClient
{
    // RPC (api.dropboxapi.com) endpoints
    private const string GetMetadataUrl        = "https://api.dropboxapi.com/2/files/get_metadata";
    private const string GetTemporaryLinkUrl   = "https://api.dropboxapi.com/2/files/get_temporary_link";
    private const string ListFolderUrl         = "https://api.dropboxapi.com/2/files/list_folder";
    private const string ListFolderContinueUrl = "https://api.dropboxapi.com/2/files/list_folder/continue";
    private const string MoveUrl               = "https://api.dropboxapi.com/2/files/move_v2";
    private const string CopyUrl               = "https://api.dropboxapi.com/2/files/copy_v2";
    private const string DeleteUrl             = "https://api.dropboxapi.com/2/files/delete_v2";
    private const string CreateFolderUrl       = "https://api.dropboxapi.com/2/files/create_folder_v2";
    // content (content.dropboxapi.com) endpoints
    private const string DownloadUrl           = "https://content.dropboxapi.com/2/files/download";
    private const string UploadUrl             = "https://content.dropboxapi.com/2/files/upload";
    private const string UploadStartUrl        = "https://content.dropboxapi.com/2/files/upload_session/start";
    private const string UploadAppendUrl       = "https://content.dropboxapi.com/2/files/upload_session/append_v2";
    private const string UploadFinishUrl       = "https://content.dropboxapi.com/2/files/upload_session/finish";

    // Dropbox single-shot upload cap is 150 MB; above that use upload_session
    // with 100 MB chunks. The /api/file/upload endpoint routes by size.
    private const long ChunkSize = 104_857_600; // 100 MB

    private readonly IDropboxTokenService _tokens;
    private readonly HttpClient _http;
    private readonly DropboxOptions _cfg;
    private readonly ILogger<DropboxApiClient> _log;

    public DropboxApiClient(
        IDropboxTokenService tokens,
        IHttpClientFactory httpFactory,
        IOptions<DropboxOptions> cfg,
        ILogger<DropboxApiClient> log)
    {
        _tokens = tokens;
        _http = httpFactory.CreateClient();
        // Default HttpClient timeout (100s) aborts multi-hundred-MB transfers;
        // raise it for the bridge->Dropbox leg (G32).
        _http.Timeout = TimeSpan.FromMinutes(30);
        _cfg = cfg.Value;
        _log = log;
    }

    // ----- read operations ---------------------------------------------------

    // get_metadata returns 409 path/not_found for a missing path. Surface that
    // as found:false (tristate) rather than throwing, so VBA's outFound works.
    public async Task<(bool Found, string? ErrorSummary, string? RawJson)> GetMetadataAsync(
        string path, CancellationToken ct = default)
    {
        using var resp = await SendRawWithRetryAsync(
            () => BuildJsonRequestAsync(GetMetadataUrl, new { path }, ct), ct);
        var body = await resp.Content.ReadAsStringAsync(ct);
        if (resp.IsSuccessStatusCode) return (true, null, body);

        var summary = TryExtractErrorSummary(body);
        if (resp.StatusCode == HttpStatusCode.Conflict &&
            (summary?.StartsWith("path/not_found", StringComparison.Ordinal) ?? false))
            return (false, summary, null);

        throw MakeError(resp.StatusCode, summary, body, resp.Headers.RetryAfter);
    }

    public async Task<string> GetTemporaryLinkAsync(string path, CancellationToken ct = default)
    {
        using var resp = await ApiPostAsync(GetTemporaryLinkUrl, new { path }, ct);
        var json = await resp.Content.ReadFromJsonAsync<JsonElement>(ct);
        return json.GetProperty("link").GetString()!;
    }

    // Returns a synthesized list_folder JSON ({entries, has_more:false, cursor:""})
    // with ALL pages concatenated — folders with >~2000 entries would otherwise
    // truncate silently (G38).
    public async Task<string> ListFolderAsync(string path, CancellationToken ct = default)
    {
        var entries = new JsonArray();
        bool hasMore;
        string? cursor;

        using (var resp = await ApiPostAsync(ListFolderUrl, new { path }, ct))
        {
            var page = await resp.Content.ReadFromJsonAsync<JsonElement>(ct);
            AppendEntries(entries, page);
            (hasMore, cursor) = ReadCursor(page);
        }
        while (hasMore)
        {
            using var resp = await ApiPostAsync(ListFolderContinueUrl, new { cursor }, ct);
            var page = await resp.Content.ReadFromJsonAsync<JsonElement>(ct);
            AppendEntries(entries, page);
            (hasMore, cursor) = ReadCursor(page);
        }

        var result = new JsonObject
        {
            ["entries"]  = entries,
            ["has_more"] = false,
            ["cursor"]   = ""
        };
        return result.ToJsonString();
    }

    // G34 → bridge proxy: stream the file bytes back to VBA. Workstations never
    // reach the Dropbox CDN directly.
    public async Task<byte[]> DownloadAsync(string path, CancellationToken ct = default)
    {
        var arg = JsonSerializer.Serialize(new { path });
        using var resp = await ContentPostAsync(DownloadUrl, arg, Array.Empty<byte>(), 0, 0, ct);
        return await resp.Content.ReadAsByteArrayAsync(ct);
    }

    // ----- write operations (gated by GuardWrites at the endpoint) -----------

    public async Task UploadAsync(string path, byte[] bytes, CancellationToken ct = default)
    {
        var arg = JsonSerializer.Serialize(new
        {
            path,
            mode = "overwrite",
            autorename = false,
            mute = false
        });
        using var _ = await ContentPostAsync(UploadUrl, arg, bytes, 0, bytes.Length, ct);
    }

    // upload_session start -> append_v2 (per 100 MB chunk) -> finish.
    public async Task UploadLargeAsync(string path, byte[] bytes, CancellationToken ct = default)
    {
        string sessionId;
        int firstLen = (int)Math.Min(ChunkSize, bytes.LongLength);
        using (var resp = await ContentPostAsync(UploadStartUrl,
                   JsonSerializer.Serialize(new { close = false }), bytes, 0, firstLen, ct))
        {
            var json = await resp.Content.ReadFromJsonAsync<JsonElement>(ct);
            sessionId = json.GetProperty("session_id").GetString()!;
        }

        long offset = firstLen;
        while (offset < bytes.LongLength)
        {
            int len = (int)Math.Min(ChunkSize, bytes.LongLength - offset);
            var arg = JsonSerializer.Serialize(new
            {
                cursor = new { session_id = sessionId, offset },
                close = false
            });
            using var resp = await ContentPostAsync(UploadAppendUrl, arg, bytes, (int)offset, len, ct);
            offset += len;
        }

        var commitArg = JsonSerializer.Serialize(new
        {
            cursor = new { session_id = sessionId, offset },
            commit = new { path, mode = "overwrite", autorename = false, mute = false }
        });
        using var _ = await ContentPostAsync(UploadFinishUrl, commitArg, Array.Empty<byte>(), 0, 0, ct);
    }

    // move_v2 / copy_v2 run autorename=false: a destination conflict comes back
    // as a Dropbox 409 to/conflict, which MakeError maps to DropboxErrorKind
    // .Conflict -> HTTP 409, letting the VBA caller tell "already exists" apart
    // from a transport/other failure (legacy behavior preserved).
    public async Task MoveAsync(string fromPath, string toPath, CancellationToken ct = default)
    {
        using var _ = await ApiPostAsync(MoveUrl, new
        {
            from_path = fromPath,
            to_path = toPath,
            allow_shared_folder = false,
            autorename = false
        }, ct);
    }

    public async Task CopyAsync(string fromPath, string toPath, CancellationToken ct = default)
    {
        using var _ = await ApiPostAsync(CopyUrl, new
        {
            from_path = fromPath,
            to_path = toPath,
            allow_shared_folder = false,
            autorename = false
        }, ct);
    }

    public async Task DeleteAsync(string path, CancellationToken ct = default)
    {
        using var _ = await ApiPostAsync(DeleteUrl, new { path }, ct);
    }

    // create_folder_v2 with autorename=false. A path/conflict ("already exists")
    // is treated as SUCCESS — the legacy VBA CreateFolder relies on this
    // idempotency and a Phase test asserts it.
    public async Task CreateFolderAsync(string path, CancellationToken ct = default)
    {
        try
        {
            using var _ = await ApiPostAsync(CreateFolderUrl, new { path, autorename = false }, ct);
        }
        catch (DropboxApiException ex) when (ex.Kind == DropboxErrorKind.Conflict)
        {
            _log.LogInformation("CreateFolder: '{Path}' already exists — treated as success.", path);
        }
    }

    // ----- transport helpers -------------------------------------------------

    private string PathRootHeader() =>
        JsonSerializer.Serialize(new Dictionary<string, string>
        {
            [".tag"]         = "namespace_id",
            ["namespace_id"] = _cfg.NamespaceId,
        });

    private async Task<HttpRequestMessage> BuildJsonRequestAsync(string url, object body, CancellationToken ct)
    {
        var token = await _tokens.GetValidAccessTokenAsync(ct);
        var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Add("Dropbox-API-Path-Root", PathRootHeader());
        req.Content = JsonContent.Create(body);
        return req;
    }

    // RPC POST that throws DropboxApiException on any non-success status.
    private async Task<HttpResponseMessage> ApiPostAsync(string url, object body, CancellationToken ct)
    {
        var resp = await SendRawWithRetryAsync(() => BuildJsonRequestAsync(url, body, ct), ct);
        if (resp.IsSuccessStatusCode) return resp;
        var ex = await BuildErrorAsync(resp, ct);
        resp.Dispose();
        throw ex;
    }

    // content-API POST with an explicit byte range as the octet-stream body.
    // Throws DropboxApiException on non-success. (Not retried — bodies may be
    // large and a chunked upload cursor cannot be safely replayed mid-session.)
    private async Task<HttpResponseMessage> ContentPostAsync(
        string url, string apiArg, byte[] data, int offset, int count, CancellationToken ct)
    {
        var token = await _tokens.GetValidAccessTokenAsync(ct);
        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Add("Dropbox-API-Path-Root", PathRootHeader());
        // JsonSerializer escapes non-ASCII to \uXXXX, keeping the header ASCII-safe.
        req.Headers.Add("Dropbox-API-Arg", apiArg);
        var content = new ByteArrayContent(data, offset, count);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
        req.Content = content;

        var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        if (resp.IsSuccessStatusCode) return resp;
        var ex = await BuildErrorAsync(resp, ct);
        resp.Dispose();
        throw ex;
    }

    // Sends the request, retrying transient failures (HTTP 429 honoring
    // Retry-After, and 5xx) with capped exponential backoff. Returns the final
    // response (success OR last failure) for the caller to interpret. The
    // request is rebuilt each attempt because content cannot be resent (G33).
    private async Task<HttpResponseMessage> SendRawWithRetryAsync(
        Func<Task<HttpRequestMessage>> requestFactory, CancellationToken ct)
    {
        const int maxAttempts = 4;
        HttpResponseMessage resp = null!;
        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            using var req = await requestFactory();
            resp = await _http.SendAsync(req, ct);
            if (resp.IsSuccessStatusCode) return resp;

            bool transient = resp.StatusCode == HttpStatusCode.TooManyRequests
                             || (int)resp.StatusCode >= 500;
            if (!transient || attempt == maxAttempts) return resp;

            int delaySec = 1 << (attempt - 1); // 1, 2, 4
            if (resp.Headers.RetryAfter?.Delta is TimeSpan ts)
                delaySec = Math.Max(delaySec, (int)ts.TotalSeconds);
            _log.LogWarning("Dropbox {Status}; retry {Attempt}/{Max} in {Delay}s.",
                (int)resp.StatusCode, attempt, maxAttempts, delaySec);
            resp.Dispose();
            await Task.Delay(TimeSpan.FromSeconds(delaySec), ct);
        }
        return resp;
    }

    // ----- error / JSON helpers ----------------------------------------------

    private async Task<DropboxApiException> BuildErrorAsync(HttpResponseMessage resp, CancellationToken ct)
    {
        string body = "";
        try { body = await resp.Content.ReadAsStringAsync(ct); } catch { /* best effort */ }
        return MakeError(resp.StatusCode, TryExtractErrorSummary(body), body, resp.Headers.RetryAfter);
    }

    private static DropboxApiException MakeError(
        HttpStatusCode status, string? summary, string body, RetryConditionHeaderValue? retryAfter)
    {
        int? retry = null;
        if (retryAfter?.Delta is TimeSpan ts) retry = (int)ts.TotalSeconds;
        else if (retryAfter?.Date is DateTimeOffset d)
            retry = Math.Max(0, (int)(d - DateTimeOffset.UtcNow).TotalSeconds);

        var kind = status switch
        {
            HttpStatusCode.Unauthorized    => DropboxErrorKind.AuthError,
            HttpStatusCode.TooManyRequests => DropboxErrorKind.RateLimited,
            HttpStatusCode.Conflict when (summary?.StartsWith("path/not_found", StringComparison.Ordinal) ?? false)
                => DropboxErrorKind.NotFound,
            HttpStatusCode.Conflict when (summary?.Contains("conflict", StringComparison.Ordinal) ?? false)
                => DropboxErrorKind.Conflict,
            _ => DropboxErrorKind.Other
        };
        var detail = string.IsNullOrEmpty(summary) ? Truncate(body, 500) : summary;
        return new DropboxApiException(kind, (int)status, summary,
            $"Dropbox API error {(int)status}: {detail}", retry);
    }

    private static string Truncate(string s, int max) => s.Length <= max ? s : s[..max];

    private static string? TryExtractErrorSummary(string body)
    {
        if (string.IsNullOrWhiteSpace(body)) return null;
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (doc.RootElement.TryGetProperty("error_summary", out var s))
                return s.GetString();
        }
        catch { /* not JSON */ }
        return null;
    }

    private static void AppendEntries(JsonArray sink, JsonElement page)
    {
        if (page.TryGetProperty("entries", out var arr) && arr.ValueKind == JsonValueKind.Array)
            foreach (var e in arr.EnumerateArray())
                sink.Add(JsonNode.Parse(e.GetRawText()));
    }

    private static (bool HasMore, string? Cursor) ReadCursor(JsonElement page)
    {
        bool hasMore = page.TryGetProperty("has_more", out var hm) && hm.GetBoolean();
        string? cursor = page.TryGetProperty("cursor", out var c) ? c.GetString() : null;
        return (hasMore, cursor);
    }
}
