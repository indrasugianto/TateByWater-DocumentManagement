using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.Server.IIS;
using Microsoft.Data.SqlClient;
using TBCMSDropboxBridge.Models;
using TBCMSDropboxBridge.Services;

var builder = WebApplication.CreateBuilder(args);

// --- Data protection --------------------------------------------------------
// Key ring persisted to a configured path AND protected at rest with
// MACHINE-scope DPAPI. Without protectToLocalMachine:true the ring is encrypted
// to the app-pool identity's profile, re-introducing exactly the profile-binding
// failure (G27/G40) this whole service exists to escape.
var dpKeyPath = builder.Configuration["DataProtection:KeyStorePath"]
                ?? @"C:\ProgramData\TBCMSBridge\dpkeys";
Directory.CreateDirectory(dpKeyPath);
builder.Services.AddDataProtection()
    .PersistKeysToFileSystem(new DirectoryInfo(dpKeyPath))
    .ProtectKeysWithDpapi(protectToLocalMachine: true)
    .SetApplicationName("TBCMSDropboxBridge");   // stable across redeploys

// --- Request body cap -------------------------------------------------------
// Default is ~30 MB, but uploads (incl. the >150 MB UploadLargeFile path) post
// raw bytes through this service. Set BOTH the in-process IIS limit (the one
// that actually applies under hostingModel=inprocess) and the Kestrel limit
// (no-op in-process, but correct if ever hosted out-of-process). The IIS
// requestFiltering maxAllowedContentLength in web.config is also required (G32).
builder.Services.Configure<IISServerOptions>(o => o.MaxRequestBodySize = 2_147_483_648); // 2 GB
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = 2_147_483_648);

// --- Options + SQL + services ----------------------------------------------
// Dev/test convenience: if no AppSecret is configured (placeholder or empty),
// pull it from tblDropboxConfig.AppSecret — the secret already lives in the DB
// on the test server, so a local `dotnet run` needs no secret in any file.
// Development ONLY; production must supply the secret via
// appsettings.Production.json / env var (the SQL copy is being rotated out).
if (builder.Environment.IsDevelopment())
{
    var configured = builder.Configuration["Dropbox:AppSecret"];
    if (string.IsNullOrEmpty(configured) ||
        configured.StartsWith("REPLACE_WITH", StringComparison.OrdinalIgnoreCase))
    {
        var fromSql = TryReadAppSecretFromSql(builder.Configuration.GetConnectionString("TateByWater"));
        if (!string.IsNullOrEmpty(fromSql))
            builder.Configuration["Dropbox:AppSecret"] = fromSql;
    }
}

builder.Services.Configure<DropboxOptions>(builder.Configuration.GetSection("Dropbox"));

builder.Services.AddScoped(_ =>
    new SqlConnection(builder.Configuration.GetConnectionString("TateByWater")));

builder.Services.AddHttpClient();
builder.Services.AddScoped<IDropboxTokenService, DropboxTokenService>();
builder.Services.AddScoped<DropboxApiClient>();

// --- Session (OAuth state for the setup flow) -------------------------------
// The setup round-trip stays on localhost (Phase E) so this cookie survives the
// Dropbox redirect.
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(o =>
{
    o.Cookie.Name = "TBCMSBridge.Setup";
    o.Cookie.HttpOnly = true;
    o.IdleTimeout = TimeSpan.FromMinutes(10);
});

// --- Windows auth -----------------------------------------------------------
// In-process IIS hosting (web.config hostingModel="inprocess") forwards Windows
// auth, so use the in-process IIS scheme (IISServerDefaults from
// Microsoft.AspNetCore.Server.IIS), NOT AddNegotiate() (Kestrel/out-of-process)
// and NOT IISDefaults (the out-of-process IISIntegration constant).
// Under IIS in-process hosting (production) use the in-process IIS scheme and
// require an authenticated Windows user on every endpoint. Under local Kestrel
// (`dotnet run`, Development) there is NO IIS handler for that scheme, so
// registering it would 401 every request; allow anonymous over localhost so the
// proxy logic can be exercised directly. Production keeps Windows auth.
if (builder.Environment.IsDevelopment())
{
    builder.Services.AddAuthentication();   // no default scheme -> UseAuthentication is a no-op
    builder.Services.AddAuthorization();    // no fallback policy -> endpoints allow anonymous
}
else
{
    builder.Services.AddAuthentication(IISServerDefaults.AuthenticationScheme);
    builder.Services.AddAuthorization(opts =>
        opts.FallbackPolicy = new AuthorizationPolicyBuilder()
            .RequireAuthenticatedUser()
            .Build());
}

var app = builder.Build();

// --- Exception -> HTTP status mapping (MUST be registered first) ------------
// Without this VBA gets a blanket 500 and can't distinguish "not configured" /
// "writes disabled" / "auth failure" / "not found" / "conflict". Maps the
// typed exceptions thrown by the services to the B.7 contract statuses.
app.UseExceptionHandler(errApp => errApp.Run(async ctx =>
{
    var ex = ctx.Features.Get<IExceptionHandlerFeature>()?.Error;

    int status = ex switch
    {
        InvalidOperationException ioe when ioe.Message.Contains("not configured")
            => StatusCodes.Status503ServiceUnavailable,
        InvalidOperationException ioe when ioe.Message.Contains("Writes are disabled")
            => StatusCodes.Status403Forbidden,
        DropboxApiException dax => dax.Kind switch
        {
            DropboxErrorKind.NotFound    => StatusCodes.Status404NotFound,
            DropboxErrorKind.Conflict    => StatusCodes.Status409Conflict,
            DropboxErrorKind.AuthError   => StatusCodes.Status401Unauthorized,
            DropboxErrorKind.RateLimited => StatusCodes.Status429TooManyRequests,
            _                            => StatusCodes.Status500InternalServerError,
        },
        HttpRequestException hre when hre.StatusCode == HttpStatusCode.Unauthorized
            => StatusCodes.Status401Unauthorized,
        _ => StatusCodes.Status500InternalServerError,
    };

    if (ex is DropboxApiException { Kind: DropboxErrorKind.RateLimited, RetryAfterSeconds: int ra })
        ctx.Response.Headers.RetryAfter = ra.ToString();

    ctx.Response.StatusCode = status;
    await ctx.Response.WriteAsJsonAsync(new { error = ex?.Message });
}));

app.UseSession();           // before endpoints; needed by the setup flow
app.UseAuthentication();
app.UseAuthorization();

// --- Server-side write guard (mirrors VBA ALLOW_DROPBOX_WRITES) -------------
// The VBA kill-switch protects nothing once the bridge exists: any domain user
// can POST a write endpoint directly. This flag is the bridge's equivalent
// boundary and MUST be false in every non-production deployment.
bool allowWrites = builder.Configuration.GetValue<bool>("Bridge:AllowWrites");
void GuardWrites()
{
    if (!allowWrites)
        throw new InvalidOperationException(
            "Writes are disabled on this bridge (Bridge:AllowWrites=false).");
}

// --- Operational endpoints (Windows Auth required) --------------------------

app.MapPost("/api/metadata", async (MetadataRequest req, DropboxApiClient db) =>
{
    var (found, errSummary, rawJson) = await db.GetMetadataAsync(req.Path);
    return Results.Ok(new MetadataResponse(found, errSummary, rawJson));
});

app.MapPost("/api/file/link", async (FileDownloadLinkRequest req, DropboxApiClient db) =>
{
    var link = await db.GetTemporaryLinkAsync(req.Path);
    return Results.Ok(new FileDownloadLinkResponse(link));
});

// G34 -> bridge proxy: stream the file bytes back so workstations never need
// outbound access to the Dropbox CDN.
app.MapPost("/api/file/download", async (FileDownloadRequest req, DropboxApiClient db) =>
{
    var bytes = await db.DownloadAsync(req.Path);
    return Results.File(bytes, "application/octet-stream");
});

app.MapPost("/api/folder/list", async (FolderListRequest req, DropboxApiClient db) =>
{
    var json = await db.ListFolderAsync(req.Path);
    return Results.Text(json, "application/json");
});

app.MapPost("/api/folder/create", async (FolderCreateRequest req, DropboxApiClient db) =>
{
    GuardWrites();
    await db.CreateFolderAsync(req.Path);   // 409 already-exists treated as success
    return Results.Ok();
});

app.MapPost("/api/file/upload", async (HttpRequest httpReq, DropboxApiClient db) =>
{
    GuardWrites();
    var path  = httpReq.Headers["X-Dropbox-Path"].ToString();
    var bytes = await BinaryBody(httpReq);
    if (bytes.LongLength > 157_286_400)   // 150 MB
        await db.UploadLargeAsync(path, bytes);
    else
        await db.UploadAsync(path, bytes);
    return Results.Ok(new UploadResponse(true, null));
});

app.MapPost("/api/file/move", async (MoveRequest req, DropboxApiClient db) =>
{
    GuardWrites();
    await db.MoveAsync(req.FromPath, req.ToPath);
    return Results.Ok();
});

app.MapPost("/api/file/copy", async (CopyRequest req, DropboxApiClient db) =>
{
    GuardWrites();
    await db.CopyAsync(req.FromPath, req.ToPath);
    return Results.Ok();
});

app.MapPost("/api/file/delete", async (DeleteRequest req, DropboxApiClient db) =>
{
    GuardWrites();
    await db.DeleteAsync(req.Path);
    return Results.Ok();
});

app.MapGet("/api/status", async (IDropboxTokenService tokens) =>
{
    if (!await tokens.HasTokenAsync())
        return Results.Ok(new StatusResponse("needs_setup", null, null));
    try
    {
        // GetValidAccessTokenAsync may return a CACHED token; VerifyLive... makes
        // a live users/get_current_account call so "ok" means genuinely live (a
        // revoked token surfaces as "error", not a stale "ok").
        var email = await tokens.VerifyLiveAccountEmailAsync();
        return Results.Ok(new StatusResponse("ok", email, null));
    }
    catch (Exception ex)
    {
        return Results.Ok(new StatusResponse("error", null, ex.Message));
    }
});

// --- Setup endpoints (restricted to loopback / localhost) -------------------

app.MapGet("/api/setup/start", (IConfiguration cfg, HttpContext ctx) =>
{
    if (!IsLocalRequest(ctx)) return Results.Forbid();
    var dropboxCfg = cfg.GetSection("Dropbox");
    var state = Guid.NewGuid().ToString("N");
    SetupState.Pending = state;   // held in-process, not the session (see SetupState)
    var authUrl = "https://www.dropbox.com/oauth2/authorize" +
        $"?client_id={dropboxCfg["AppKey"]}" +
        "&response_type=code&token_access_type=offline" +
        $"&state={state}" +
        $"&redirect_uri={Uri.EscapeDataString(dropboxCfg["RedirectUri"]!)}";
    return Results.Redirect(authUrl);
});

app.MapGet("/api/setup/callback", async (
    string? code, string? state, string? error,
    HttpContext ctx,
    IDropboxTokenService tokens,
    IConfiguration cfg) =>
{
    if (!IsLocalRequest(ctx)) return Results.Forbid();
    // Dropbox sends ?error=access_denied (no code) if the admin clicks Deny.
    if (!string.IsNullOrEmpty(error))
        return Results.BadRequest($"Dropbox authorization was declined: {error}");
    if (string.IsNullOrEmpty(code))
        return Results.BadRequest("Missing authorization code.");
    // CSRF state check WITHOUT the in-memory session, which proved fragile across
    // the external Dropbox redirect on locked-down server browsers. The state is
    // held in a process-static set by /api/setup/start (single worker process).
    // Both setup endpoints are loopback-only (IsLocalRequest) and require an
    // authenticated admin, so this is a sufficient CSRF guard. If the static was
    // cleared (e.g., an app-pool recycle between start and callback) don't
    // hard-fail — the loopback + Windows-auth gates still apply.
    var storedState = SetupState.Pending;
    if (!string.IsNullOrEmpty(storedState) && state != storedState)
        return Results.BadRequest("State mismatch");
    SetupState.Pending = null;

    var dropboxCfg = cfg.GetSection("Dropbox");
    using var http = new HttpClient();
    var form = new Dictionary<string, string>
    {
        ["code"]          = code,
        ["grant_type"]    = "authorization_code",
        ["client_id"]     = dropboxCfg["AppKey"]!,
        ["client_secret"] = dropboxCfg["AppSecret"]!,
        ["redirect_uri"]  = dropboxCfg["RedirectUri"]!,
    };
    using var resp = await http.PostAsync(
        "https://api.dropboxapi.com/oauth2/token",
        new FormUrlEncodedContent(form));
    resp.EnsureSuccessStatusCode();
    var json      = await resp.Content.ReadFromJsonAsync<JsonElement>();
    var access    = json.GetProperty("access_token").GetString()!;
    var refresh   = json.GetProperty("refresh_token").GetString()!;
    var expiresIn = json.GetProperty("expires_in").GetInt32();

    // Capture the account email now — the only place with a fresh, known-good
    // token; surfaced by /api/status.
    string accountEmail = "";
    using (var who = new HttpClient())
    {
        who.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", access);
        using var acctResp = await who.PostAsync(
            "https://api.dropboxapi.com/2/users/get_current_account", null);
        if (acctResp.IsSuccessStatusCode)
        {
            var acct = await acctResp.Content.ReadFromJsonAsync<JsonElement>();
            accountEmail = acct.GetProperty("email").GetString() ?? "";
        }
    }

    await tokens.SaveTokensAsync(access, refresh, expiresIn, accountEmail,
                                 ctx.User.Identity?.Name ?? "setup", default);
    return Results.Text("Setup complete. You can close this browser tab.");
});

app.Run();

// --- Helpers ----------------------------------------------------------------

static bool IsLocalRequest(HttpContext ctx)
{
    var remote = ctx.Connection.RemoteIpAddress;
    return remote != null && (
        IPAddress.IsLoopback(remote) ||
        remote.Equals(ctx.Connection.LocalIpAddress));
}

static async Task<byte[]> BinaryBody(HttpRequest req)
{
    using var ms = new MemoryStream();
    await req.Body.CopyToAsync(ms);
    return ms.ToArray();
}

// Development-only fallback: read the existing Dropbox AppSecret from
// tblDropboxConfig so a local `dotnet run` needs no secret in any file.
static string? TryReadAppSecretFromSql(string? connStr)
{
    if (string.IsNullOrEmpty(connStr)) return null;
    try
    {
        using var cn = new SqlConnection(connStr);
        cn.Open();
        using var cmd = new SqlCommand(
            "SELECT AppSecret FROM dbo.tblDropboxConfig WHERE ConfigID = 1", cn);
        return cmd.ExecuteScalar() as string is { Length: > 0 } s ? s : null;
    }
    catch { return null; }
}

// Holds the one-time setup OAuth CSRF state across the start->callback redirect,
// independent of the in-memory session/cookie (which didn't survive the external
// redirect on the locked-down server browser). Safe because the app pool runs a
// single worker process and the setup endpoints are loopback-only + admin-authed.
static class SetupState
{
    public static volatile string? Pending;
}
