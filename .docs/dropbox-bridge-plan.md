# TBCMS Dropbox Bridge Service — Implementation Plan

> **Purpose of this document:** Step-by-step guide for implementing a lightweight
> internal REST service that proxies all Dropbox API calls on behalf of the MS
> Access VBA frontend.  This supersedes the VBA-direct OAuth approach that was
> causing authentication failures on locked-down workstations (G27).
>
> **Intended audience:** A fresh Claude Code session or a developer who has read
> `CLAUDE.md` but has not worked on the project before.  Assumes `CLAUDE.md` has
> already been read for project context.
>
> **Branch to work on:** branch from `main`. (The earlier
> `claude/ms-access-dropbox-auth-udqa7w` branch was already merged via PR #1
> — do **not** keep working on it.)

> ## ▶ Implementation status (2026-06-22)
>
> Phases **A–C + F** are implemented on branch `claude/dropbox-bridge`:
> - **Phase A** — Section 9 appended to the SQL installer (`BridgeUrl` column +
>   `tblDropboxServiceToken`), idempotent. *Not yet applied to the dev DB —
>   `sqlcmd` was unavailable in the implementing session; run the installer
>   (or just Section 9) against `awsql2022dev` as a deploy step.*
> - **Phase B** — `dropbox-bridge/` ASP.NET Core 8 project builds clean
>   (`dotnet build`, 0 warnings). Includes the extra **`/api/file/download`**
>   proxy endpoint (see G34 decision below), 429/Retry-After handling (G33),
>   `list_folder` pagination (G38), and Dropbox→HTTP error translation (B.7).
> - **Phase C** — `DropboxService.bas` rewired to the bridge. The pre-bridge
>   OAuth/DPAPI/token code is preserved in-place behind the module directive
>   `#Const PREBRIDGE_LEGACY = False`; flipping it to `True` restores the
>   original behaviour (this is the bridge-era form of the project's
>   "comment-out the LEGACY block" rollback convention). `OpenDocument` uses
>   the bridge proxy; `EnsureValidToken` is a True-returning shim (still called
>   by `DocumentManagement.bas`).
> - **Phase F** — `Phase3a`/`Phase3e` smoke tests rewritten for the bridge;
>   `PhaseBridge_ConnectivityTest` added; legacy Phase 3b tests retired behind
>   the directive.
>
> **Not done (deployment / hardware-dependent):** Phases **D (IIS deploy)** and
> **E (one-time OAuth setup)** are manual server actions; **VBA was not
> compiled** (no MS Access in the implementing session — import into
> `TBCMS_Test.accde` and run Phase F). The real `AppSecret` and the production
> SQL host still need to go into `appsettings.Production.json` on the server.

---

## Why we're doing this (problem recap)

The existing `DropboxService.bas` VBA module authenticates directly with the
Dropbox API from each user's MS Access session.  This design has three structural
problems that are hard to fix in VBA:

1. **DPAPI binding.** Tokens in `tblDropboxTokens` are encrypted with Windows
   DPAPI, which binds ciphertext to the individual Windows user profile.  Every
   new staff member, every reimaged PC, or any cross-profile scenario triggers a
   full re-authentication.  This already caused a confirmed production failure
   (see G27 in `.docs/dropbox-migration-plan.md`).

2. **OAuth callback capture.** The only working fallback on locked-down
   workstations is the manual-paste `InputBox` flow: the browser shows
   "site can't be reached", and the user must copy the full redirect URL from the
   address bar and paste it into an Access dialog.  Non-technical users will fail
   and call IT.  The automatic `HttpListener` path requires PowerShell and admin
   rights — both blocked on this firm's machines.

3. **AppSecret client-side.** The Dropbox `AppSecret` is loaded from SQL Server
   into VBA memory at startup.  It should never leave the server.

The fix: a thin internal REST service runs on the firm's server and owns the
Dropbox OAuth tokens (one service account, stored server-side).  VBA calls the
service via Windows Integrated Auth — zero new credentials, zero OAuth prompts
for staff.

---

## Solution architecture

```
MS Access VBA (DropboxService.bas)
        │
        │  WinHttp, Windows Integrated Auth (NTLM)
        │  Simple JSON request/response, no Dropbox headers
        ▼
TBCMSDropboxBridge   (ASP.NET Core 8 Minimal API, hosted on IIS)
        │
        │  HTTPS, Bearer token, Dropbox-API-Path-Root header
        │  OAuth refresh handled automatically
        ▼
Dropbox Business API  (api.dropboxapi.com / content.dropboxapi.com)
```

**Key properties:**
- VBA holds **no** secrets (no AppKey, no AppSecret, no OAuth tokens, no DPAPI)
- One Dropbox service-account token stored encrypted in SQL Server, accessible
  only to the bridge service
- Initial authentication is a one-time admin action via a setup page; after that
  the service refreshes tokens automatically and VBA never sees the OAuth flow
- All existing public VBA function signatures are **unchanged** — the 8 calling
  forms need zero edits
- Windows Integrated Auth between VBA and the bridge means staff never enter a
  password; their domain login is sufficient

---

## Repository layout (what gets added)

```
TateByWater-CMS/
├── dropbox-bridge/                      ← NEW: the .NET project lives here
│   ├── TBCMSDropboxBridge.csproj
│   ├── appsettings.json                 ← AppKey, AppSecret, SQL conn string
│   ├── appsettings.Production.json      ← override for IIS deployment
│   ├── Program.cs                       ← app wiring + all routes
│   ├── Services/
│   │   ├── DropboxTokenService.cs       ← loads/saves/refreshes service tokens
│   │   └── DropboxApiClient.cs          ← thin wrapper over all Dropbox API calls
│   └── Models/
│       ├── BridgeRequest.cs             ← shared request/response POCOs
│       └── DropboxModels.cs             ← Dropbox-specific response shapes
├── Dropbox-Migration/
│   ├── DropboxService.bas               ← MODIFIED: OAuth stripped, BridgeRequest added
│   ├── DocumentManagement.bas           ← UNCHANGED (public signatures unchanged)
│   └── Dropbox-Migration-SQL-Install.sql ← MODIFIED: adds 2 SQL objects
└── .docs/
    └── dropbox-bridge-plan.md           ← THIS FILE
```

---

## Phase A — SQL schema additions

Add two objects to the SQL installer.  Append them to
`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql` as a new **Section 9**
at the end of the file.  Both are idempotent (IF NOT EXISTS guarded).

### A.1 — `BridgeUrl` column on `tblDropboxConfig`

VBA reads the bridge service URL from here at startup, so it can be changed
without recompiling the `.accde`.

```sql
-- Section 9.1 — Add BridgeUrl to tblDropboxConfig
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.tblDropboxConfig')
      AND name = 'BridgeUrl'
)
BEGIN
    ALTER TABLE dbo.tblDropboxConfig
        ADD BridgeUrl nvarchar(500) NULL;

    UPDATE dbo.tblDropboxConfig
    SET    BridgeUrl = N'http://tbcms-bridge.tatebywater.local/api'
    WHERE  ConfigID = 1;

    PRINT 'Section 9.1: BridgeUrl column added and seeded.';
END
ELSE
    PRINT 'Section 9.1: BridgeUrl already present — skipped.';
```

After running the installer, IT must update the URL to match wherever the IIS
site is actually hosted:

```sql
UPDATE dbo.tblDropboxConfig
SET    BridgeUrl = N'http://<server-name-or-ip>/tbcms-bridge/api'
WHERE  ConfigID = 1;
```

> **Re-run caveat.** `tblDropboxConfig` is `DROP`/`CREATE`d by the installer's
> Section 1.2, so a full destructive re-run drops the `BridgeUrl` column; Section
> 9.1 (running later in the same pass) re-adds and re-seeds it to the placeholder
> URL — IT must re-apply the real URL afterward, exactly as for `AppSecret`.
> (Cleaner long-term: fold `BridgeUrl` into the Section 1.2 `CREATE TABLE` so the
> column has one canonical definition; the Section 9 `ALTER` works but splits the
> schema across two sections.)

### A.2 — `tblDropboxServiceToken` table

Holds the single service-account token row used by the bridge.  The bridge reads
and writes this table; VBA never touches it.  Tokens are stored encrypted via
.NET Data Protection (machine-scope DPAPI, not user-scope — no profile binding).

```sql
-- Section 9.2 — tblDropboxServiceToken
IF NOT EXISTS (
    SELECT 1 FROM sys.tables
    WHERE object_id = OBJECT_ID('dbo.tblDropboxServiceToken')
)
BEGIN
    CREATE TABLE dbo.tblDropboxServiceToken (
        TokenID        int           NOT NULL PRIMARY KEY,   -- always 1 (singleton)
        AccessToken    nvarchar(max) NOT NULL,   -- Data-Protection-encrypted (machine scope)
        RefreshToken   nvarchar(max) NOT NULL,   -- Data-Protection-encrypted (machine scope)
        ExpiresAtUtc   datetime2     NOT NULL,
        AccountEmail   nvarchar(200) NULL,
        UpdatedAtUtc   datetime2     NOT NULL DEFAULT GETUTCDATE(),
        SetupByUser    nvarchar(200) NULL,       -- Windows login that ran setup
        -- Single-row guarantee — mirrors tblDropboxConfig's CK_..._SingleRow.
        -- The bridge UPSERTs TokenID = 1; it must never accumulate rows.
        CONSTRAINT CK_DropboxServiceToken_SingleRow CHECK (TokenID = 1)
    );

    PRINT 'Section 9.2: tblDropboxServiceToken created.';
END
ELSE
    PRINT 'Section 9.2: tblDropboxServiceToken already present — skipped.';
```

---

## Phase B — Bridge service build

Create the project at `dropbox-bridge/`.  All commands below are run from the
repo root.

### B.1 — Project scaffold

```bash
dotnet new webapi -n TBCMSDropboxBridge -o dropbox-bridge --framework net8.0
cd dropbox-bridge
dotnet add package Microsoft.Data.SqlClient        # SQL Server access
```

Delete the generated `WeatherForecast.cs` and `Controllers/WeatherForecastController.cs`.

> **Package notes:**
> - `Microsoft.AspNetCore.DataProtection` is part of the ASP.NET Core **shared
>   framework** for `net8.0` — no `dotnet add package` needed (the earlier draft
>   listed it; harmless but unnecessary).
> - Because Windows auth uses the in-process **IIS** scheme (`IISServerDefaults`,
>   see Program.cs B.6), the `Microsoft.AspNetCore.Authentication.Negotiate`
>   package is **not** required. Only add Negotiate if you switch to
>   out-of-process/Kestrel hosting.

### B.2 — `appsettings.json`

```json
{
  "Logging": {
    "LogLevel": { "Default": "Information", "Microsoft.AspNetCore": "Warning" }
  },
  "AllowedHosts": "*",
  "Dropbox": {
    "AppKey":      "dqleswbnux8k3m5",
    "AppSecret":   "REPLACE_WITH_REAL_SECRET",
    "NamespaceId": "14334595683",
    "RedirectUri": "http://localhost/api/setup/callback"
  },
  "ConnectionStrings": {
    "TateByWater": "Server=awsql2022dev;Database=TateByWater;Integrated Security=True;TrustServerCertificate=True;"
  },
  "Bridge": {
    "AllowWrites": false,
    "AllowedAdGroups": [ "Domain\\TBCMSUsers", "Domain\\IT" ],
    "SetupAllowedFrom": [ "127.0.0.1", "::1" ]
  },
  "DataProtection": {
    "KeyStorePath": "C:\\ProgramData\\TBCMSBridge\\dpkeys"
  }
}
```

> **`Bridge:AllowWrites`** is the server-side equivalent of the VBA
> `ALLOW_DROPBOX_WRITES` kill-switch (see Program.cs `GuardWrites`). Keep it
> `false` in the committed `appsettings.json` and in every **test** deployment —
> the test environment treats `/Company` as read-only, and the VBA flag alone
> does NOT protect Dropbox once the bridge exists (any domain user can POST a
> write endpoint directly). Set it `true` only in the production deployment's
> `appsettings.Production.json`.

> **`Bridge:AllowedAdGroups` and `Bridge:SetupAllowedFrom` are reserved — not
> yet enforced by any code.** As written, `Program.cs` authorizes with a bare
> `RequireAuthenticatedUser()` (no group check) and the setup endpoints gate on
> a hardcoded loopback check (`IsLocalRequest`), not `SetupAllowedFrom`. These
> two keys are placeholders for access-scoping decisions still open in **G28**
> (AD group names) and **G37** (per-path / per-operation authorization). Either
> wire them up or treat them as documentation — do **not** assume they restrict
> anything today.

> **`RedirectUri` must exactly match** (a) the route (`/api/setup/callback`) and
> (b) what is registered in the Dropbox App Console. The whole setup round-trip
> runs on `localhost` (Phase E) so this is `http://localhost/api/setup/callback`
> — Dropbox permits `http://localhost` redirect URIs but **rejects non-localhost
> `http://` URIs** (they must be `https://`). See G30.

> **Connection string is per-environment.** This dev value (`awsql2022dev` /
> `TateByWater`) is correct for the **test** bridge only. The production bridge
> must point at the production SQL host (`tbf-cms`) in its
> `appsettings.Production.json`, and `tblDropboxServiceToken` + its token row
> must be provisioned **separately** in each database. A production bridge
> pointed at the dev DB would violate the "two environments, never mix" rule
> (CLAUDE.md). See G35.

> **Security note:** `AppSecret` must be set in `appsettings.Production.json`
> (which is gitignored) or via an environment variable
> `Dropbox__AppSecret=<value>` on the IIS server.  Never commit the real secret.
> Add `appsettings.Production.json` to `.gitignore`.

### B.3 — `Models/BridgeRequest.cs`

```csharp
namespace TBCMSDropboxBridge.Models;

public record MetadataRequest(string Path);
public record MetadataResponse(bool Found, string? ErrorSummary, string? RawJson);

public record FileDownloadLinkRequest(string Path);
public record FileDownloadLinkResponse(string TemporaryLink);

public record FolderListRequest(string Path);
// response is raw Dropbox JSON string

public record FolderCreateRequest(string Path);
public record MoveRequest(string FromPath, string ToPath);
public record CopyRequest(string FromPath, string ToPath);
public record DeleteRequest(string Path);

// Upload: path supplied via X-Dropbox-Path header; body = raw file bytes
public record UploadResponse(bool Success, string? ErrorDetail);

public record StatusResponse(string Status, string? AccountEmail, string? ErrorDetail);
```

### B.4 — `Services/DropboxTokenService.cs`

Responsible for: loading tokens from `tblDropboxServiceToken`, refreshing when
expiring, saving new tokens after OAuth exchange or refresh.

Key methods:

```csharp
public interface IDropboxTokenService
{
    Task<string?> GetValidAccessTokenAsync(CancellationToken ct = default);
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
```

Implementation notes:
- Use `IDataProtector` (injected via `IDataProtectionProvider`) with purpose
  string `"TBCMSDropboxBridge.ServiceToken"`. Encryption-at-rest scope is set in
  `Program.cs` via `ProtectKeysWithDpapi(protectToLocalMachine: true)` — do not
  rely on the per-user default (see G27 / the Program.cs note).
- **Single-row semantics.** `tblDropboxServiceToken` must hold exactly one row.
  `SaveTokensAsync` is an **UPSERT** (update the existing row if present, else
  insert), and every read is `SELECT TOP 1 ... ORDER BY UpdatedAtUtc DESC`. The
  schema (Section 9.2) enforces this with a single-row CHECK constraint. Do not
  `INSERT` a new row on every refresh.
- On `GetValidAccessTokenAsync`: load the row, decrypt `AccessToken`. If
  `ExpiresAtUtc - 5 minutes < UtcNow`, call `RefreshAsync` first, then decrypt
  the newly saved `AccessToken`. **Serialize concurrent refreshes** with a
  `SemaphoreSlim` around the refresh+save so simultaneous near-expiry requests
  don't double-refresh / race on the SQL write. The semaphore **must be
  `static`** (or held by a singleton): `DropboxTokenService` is registered
  **scoped** (Program.cs), so a per-instance `SemaphoreSlim` field gives every
  request its own instance and serializes nothing.
- `RefreshAsync` calls `https://api.dropboxapi.com/oauth2/token` with
  `grant_type=refresh_token` and saves the result via `SaveTokensAsync`. It MUST
  preserve the existing `AccountEmail` (load it from the current row) — do not
  pass `""`, which would wipe the stored email on every refresh.
- `GetAccountEmailAsync` returns the stored `AccountEmail` with **no** network
  call — this is what `RefreshAsync` uses to preserve the email across a refresh.
  (A live call there is impossible: refresh runs *precisely because* the token is
  expired, so it could not authenticate a `users/get_current_account` request.)
  The `/api/status` liveness guarantee instead uses the **separate**
  `VerifyLiveAccountEmailAsync`, which makes a fresh `users/get_current_account`
  call so a revoked token surfaces as an error rather than a stale "ok". Keep
  these two responsibilities in separate methods — do **not** overload one.
- Wrap the SQL connection in the injected `SqlConnection` (registered as scoped
  with the connection string from config).
- Raise `InvalidOperationException("Bridge not configured — run setup first")`
  if the token table is empty.

Full `RefreshAsync` body:

```csharp
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
    var resp = await http.PostAsync(
        "https://api.dropboxapi.com/oauth2/token",
        new FormUrlEncodedContent(form), ct);
    resp.EnsureSuccessStatusCode();
    var json = await resp.Content.ReadFromJsonAsync<JsonElement>(ct);
    var newAccess    = json.GetProperty("access_token").GetString()!;
    var newExpiresIn = json.GetProperty("expires_in").GetInt32();
    // keep existing refresh token (Dropbox does not rotate refresh tokens for
    // offline_access grants) AND preserve the stored account email — passing ""
    // here wipes it on every refresh.
    var existingEmail = await GetAccountEmailAsync(ct) ?? "";
    await SaveTokensAsync(newAccess, refreshToken, newExpiresIn,
                          existingEmail, "auto-refresh", ct);
}
```

### B.5 — `Services/DropboxApiClient.cs`

Thin wrapper over the Dropbox API.  Inject `IDropboxTokenService` (`_tokens`),
an `HttpClient` (`_http`, from `IHttpClientFactory`), and the bound `Dropbox`
options (`_cfg`). The samples below use `_cfg.AppKey` / `_cfg.AppSecret` /
`_cfg.NamespaceId`, so register the options once:
`builder.Services.Configure<DropboxOptions>(builder.Configuration.GetSection("Dropbox"))`
and inject `IOptions<DropboxOptions>` (this binding is assumed but never shown in
Program.cs — add it).

Key pattern — all API helpers follow this shape:

```csharp
// Build the path-root header value WITHOUT hand-escaping — serialize a real
// object so the JSON is guaranteed valid. The required shape is exactly:
//   {".tag":"namespace_id","namespace_id":"14334595683"}
// (matches DropboxPathRootHeader() in DropboxService.bas). Note the key is
// ".tag" — earlier drafts of this plan had a mis-escaped ".\.tag" that Dropbox
// rejects.
private string PathRootHeader() =>
    JsonSerializer.Serialize(new Dictionary<string, string>
    {
        [".tag"]        = "namespace_id",
        ["namespace_id"] = _cfg.NamespaceId,
    });

private async Task<HttpResponseMessage> ApiPostAsync(
    string url, object body, CancellationToken ct)
{
    var token = await _tokens.GetValidAccessTokenAsync(ct);
    using var req = new HttpRequestMessage(HttpMethod.Post, url);
    req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    req.Headers.Add("Dropbox-API-Path-Root", PathRootHeader());
    req.Content = JsonContent.Create(body);
    return await _http.SendAsync(req, ct);
}
```

> **Note — uploads do not use `ApiPostAsync`.** The content API
> (`content.dropboxapi.com/2/files/upload`) needs `Content-Type:
> application/octet-stream`, the path supplied via a JSON-encoded
> `Dropbox-API-Arg` header (not the body), **and** the same
> `Dropbox-API-Path-Root` header. `UploadAsync` / `UploadLargeAsync` must build
> their own `HttpRequestMessage` — reuse `PathRootHeader()` but set the raw byte
> content and `Dropbox-API-Arg` explicitly.

Methods to implement (one per public VBA function being replaced):

| Method | Dropbox endpoint | Notes |
|---|---|---|
| `GetMetadataAsync(path)` | `POST /2/files/get_metadata` | Returns `(found, errorSummary, rawJson)` |
| `GetTemporaryLinkAsync(path)` | `POST /2/files/get_temporary_link` | Returns the `link` string |
| `ListFolderAsync(path)` | `POST /2/files/list_folder` | Returns raw JSON string |
| `UploadAsync(path, bytes)` | `POST /2/files/upload` (content API) | Single-shot ≤150 MB |
| `UploadLargeAsync(path, bytes)` | upload_session/{start,append_v2,finish} | Chunks at 100 MB |
| `MoveAsync(from, to)` | `POST /2/files/move_v2` | |
| `CopyAsync(from, to)` | `POST /2/files/copy_v2` | |
| `DeleteAsync(path)` | `POST /2/files/delete_v2` | |
| `CreateFolderAsync(path)` | `POST /2/files/create_folder_v2` | `autorename: false` |

Every method must:
1. Inject the `Dropbox-API-Path-Root` namespace header (same namespace ID as
   in `tblDropboxRootConfig`, pulled from `appsettings.json`).
2. Call `_tokens.GetValidAccessTokenAsync()` — the token service handles
   refresh; callers never see raw tokens.
3. For upload: use `content.dropboxapi.com/2/files/upload`; supply the path via
   the `Dropbox-API-Arg` header (JSON-encoded), not in the URL.

> **NamespaceId now has two homes.** The bridge reads it from `appsettings.json`
> (`Dropbox:NamespaceId`); SQL `tblDropboxConfig` / `tblDropboxRootConfig` still
> carry it, but VBA no longer reads it post-rewrite. Treat `appsettings.json` as
> authoritative for the bridge, and keep the SQL copy in sync if the tenant
> namespace ever changes.

### B.6 — `Program.cs` — route registration

```csharp
var builder = WebApplication.CreateBuilder(args);

// Data protection — key ring persisted to configured path AND protected at rest
// with MACHINE-scope DPAPI. Without protectToLocalMachine:true the ring is
// encrypted to the app-pool identity's profile, which re-introduces exactly the
// profile-binding failure (G27) this whole service exists to escape: change the
// app-pool identity or lose its profile and the stored tokens become
// undecryptable, forcing a full re-setup.
builder.Services.AddDataProtection()
    .PersistKeysToFileSystem(
        new DirectoryInfo(builder.Configuration["DataProtection:KeyStorePath"]!))
    .ProtectKeysWithDpapi(protectToLocalMachine: true)
    .SetApplicationName("TBCMSDropboxBridge");   // stable across redeploys

// Raise the request-body cap — default is ~30 MB, but uploads (incl. the
// >150 MB UploadLargeFile path) post raw file bytes through this service.
// Set BOTH the Kestrel/generic limit and the IIS in-process limit.
builder.Services.Configure<IISServerOptions>(o => o.MaxRequestBodySize = 2_147_483_648); // 2 GB
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = 2_147_483_648);

// SQL client
builder.Services.AddScoped(_ =>
    new SqlConnection(
        builder.Configuration.GetConnectionString("TateByWater")));

// App services
builder.Services.AddHttpClient();
builder.Services.AddScoped<IDropboxTokenService, DropboxTokenService>();
builder.Services.AddScoped<DropboxApiClient>();

// Session — REQUIRED for the OAuth state param in the setup flow.
// (Earlier drafts called a non-existent app.AddSession(); these are the real
// registrations.) Note the setup flow keeps the entire round-trip on localhost
// so the session cookie survives the Dropbox redirect — see Phase E.
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(o =>
{
    o.Cookie.Name = "TBCMSBridge.Setup";
    o.Cookie.HttpOnly = true;
    o.IdleTimeout = TimeSpan.FromMinutes(10);
});

// Windows auth. IMPORTANT: D.2 specifies in-process IIS hosting
// (hostingModel="inprocess"), so Windows auth is forwarded by IIS — use the
// in-process IIS scheme, NOT AddNegotiate() (that is for Kestrel /
// out-of-process only). Use IISServerDefaults (from
// Microsoft.AspNetCore.Server.IIS, the in-process server) — NOT IISDefaults,
// which is the out-of-process IISIntegration constant. Both happen to resolve
// to the string "Windows", but match the constant to the hosting model.
builder.Services.AddAuthentication(IISServerDefaults.AuthenticationScheme);
builder.Services.AddAuthorization(opts =>
    opts.FallbackPolicy = new AuthorizationPolicyBuilder()
        .RequireAuthenticatedUser()
        .Build());

var app = builder.Build();

// Exception → HTTP status mapping. MUST be registered FIRST — before the
// middleware and routes below — so VBA receives 503 / 403 / 401 instead of a
// blanket 500. Full handler body is in B.7; paste it here.
app.UseExceptionHandler(/* see B.7 — maps InvalidOperationException / 401 / etc. */);

app.UseSession();           // before endpoints; needed by setup flow
app.UseAuthentication();
app.UseAuthorization();

// --- Server-side write guard (mirrors VBA ALLOW_DROPBOX_WRITES) -------------
// The VBA kill-switch protects nothing once the bridge exists: any domain user
// can POST /api/file/delete directly. This flag is the bridge's equivalent
// boundary and MUST be false in every non-production deployment (test env
// treats /Company as read-only). Read from config: "Bridge:AllowWrites".
bool allowWrites = builder.Configuration.GetValue<bool>("Bridge:AllowWrites");
void GuardWrites()
{
    if (!allowWrites)
        throw new InvalidOperationException(
            "Writes are disabled on this bridge (Bridge:AllowWrites=false).");
}

// --- Operational endpoints (Windows Auth required) -------------------------

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

app.MapPost("/api/folder/list", async (FolderListRequest req, DropboxApiClient db) =>
{
    var json = await db.ListFolderAsync(req.Path);
    return Results.Text(json, "application/json");
});

app.MapPost("/api/folder/create", async (FolderCreateRequest req, DropboxApiClient db) =>
{
    GuardWrites();
    // CreateFolderAsync MUST treat Dropbox 409 path/conflict/folder
    // ("already exists") as success — the legacy VBA CreateFolder relies on
    // this idempotency (DropboxService.bas ~2889) and a Phase test asserts it.
    await db.CreateFolderAsync(req.Path);
    return Results.Ok();
});

app.MapPost("/api/file/upload", async (HttpRequest httpReq, DropboxApiClient db) =>
{
    GuardWrites();
    var path  = httpReq.Headers["X-Dropbox-Path"].ToString();
    var bytes = await BinaryBody(httpReq);
    if (bytes.Length > 157_286_400)   // 150 MB
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
        // GetValidAccessTokenAsync may return a CACHED token without contacting
        // Dropbox, so "ok" alone does not prove the token is still valid (it
        // could have been revoked admin-side). VerifyLiveAccountEmailAsync makes
        // a cheap users/get_current_account call so "ok" means genuinely live.
        var email = await tokens.VerifyLiveAccountEmailAsync();
        return Results.Ok(new StatusResponse("ok", email, null));
    }
    catch (Exception ex)
    {
        return Results.Ok(new StatusResponse("error", null, ex.Message));
    }
});

// --- Setup endpoints (restricted to loopback / localhost) ------------------

app.MapGet("/api/setup/start", (IConfiguration cfg, HttpContext ctx) =>
{
    // Loopback-only is enforced in-code (IsLocalRequest). There is no
    // "LocalOnly" auth policy — earlier drafts referenced one that was never
    // defined and would have thrown at startup.
    if (!IsLocalRequest(ctx)) return Results.Forbid();
    var dropboxCfg = cfg.GetSection("Dropbox");
    var state = Guid.NewGuid().ToString("N");
    ctx.Session.SetString("oauth_state", state);
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
    // Dropbox sends ?error=access_denied (and NO code) if the admin clicks Deny
    // on the consent screen. Handle it explicitly — otherwise the missing `code`
    // makes minimal-API binding fail with an opaque 400.
    if (!string.IsNullOrEmpty(error))
        return Results.BadRequest($"Dropbox authorization was declined: {error}");
    if (string.IsNullOrEmpty(code))
        return Results.BadRequest("Missing authorization code.");
    var storedState = ctx.Session.GetString("oauth_state");
    if (state != storedState) return Results.BadRequest("State mismatch");

    var dropboxCfg = cfg.GetSection("Dropbox");
    using var http  = new HttpClient();
    var form = new Dictionary<string, string>
    {
        ["code"]          = code,
        ["grant_type"]    = "authorization_code",
        ["client_id"]     = dropboxCfg["AppKey"]!,
        ["client_secret"] = dropboxCfg["AppSecret"]!,
        ["redirect_uri"]  = dropboxCfg["RedirectUri"]!,
    };
    var resp = await http.PostAsync(
        "https://api.dropboxapi.com/oauth2/token",
        new FormUrlEncodedContent(form));
    resp.EnsureSuccessStatusCode();
    var json       = await resp.Content.ReadFromJsonAsync<JsonElement>();
    var access     = json.GetProperty("access_token").GetString()!;
    var refresh    = json.GetProperty("refresh_token").GetString()!;
    var expiresIn  = json.GetProperty("expires_in").GetInt32();

    // Capture the account email now (the only place we have a fresh token in a
    // known-good state). It is surfaced by /api/status and the acceptance
    // criteria expect it; the previous "" placeholder left it permanently null.
    string accountEmail = "";
    using (var who = new HttpClient())
    {
        who.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", access);
        var acctResp = await who.PostAsync(
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
```

Helper at the bottom of `Program.cs`:

```csharp
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
```

### B.7 — Error handling convention

All endpoints should catch `InvalidOperationException` (bridge not configured),
Dropbox 409 (path not found, conflict) and HTTP 5xx, and return appropriate
HTTP status codes that VBA can test:

| Condition | HTTP status | VBA handling |
|---|---|---|
| Success | 200 | Parse response |
| Bridge not yet set up | 503 | Surface "Contact IT" message |
| Writes disabled on this bridge | 403 | Surface "writes disabled" (test env) |
| Dropbox path not found | 404 | Treated as `outFound = False` |
| Dropbox auth failure | 401 | Log + show "Dropbox auth error — contact IT" |
| Any other error | 500 | Log detail; surface generic error |

> **Note on `/api/metadata`:** path-not-found is NOT surfaced as HTTP 404 — the
> endpoint returns `200` with `found:false` so VBA's `outFound` tristate works.
> The 404 row above applies to other endpoints that genuinely cannot proceed
> without the path.

> **The 404 / 401 rows are a contract the shown code does not yet fulfill.** The
> handler below only produces 503 / 403 / 401 / 500, and the `DropboxApiClient`
> methods (B.5) don't translate Dropbox errors. For **401** to fire, each method
> must call `resp.EnsureSuccessStatusCode()` (which populates
> `HttpRequestException.StatusCode`). And Dropbox returns **HTTP 409** with a
> `path/not_found` `error_summary` — *not* 404 — for missing paths, so "path not
> found → 404" requires the client to parse the error body and throw a
> 404-mapped exception. Implement that translation in `DropboxApiClient`, or the
> table's finer-grained statuses never reach VBA (everything non-503/403 lands
> on 500).

Register the exception-handling middleware **immediately after `app.Build()`**,
before `UseAuthentication`/route registration, and map exception types to the
statuses above (the previous version collapsed everything to 500, so VBA could
never distinguish "not configured" / "writes disabled" / "auth failure"):

```csharp
app.UseExceptionHandler(errApp => errApp.Run(async ctx =>
{
    var ex = ctx.Features.Get<IExceptionHandlerFeature>()?.Error;
    ctx.Response.StatusCode = ex switch
    {
        // "Bridge not configured — run setup first"
        InvalidOperationException ioe when ioe.Message.Contains("not configured")
            => StatusCodes.Status503ServiceUnavailable,
        // "Writes are disabled on this bridge" (GuardWrites)
        InvalidOperationException ioe when ioe.Message.Contains("Writes are disabled")
            => StatusCodes.Status403Forbidden,
        // Dropbox auth failures bubble up as HttpRequestException w/ 401
        HttpRequestException hre when hre.StatusCode == HttpStatusCode.Unauthorized
            => StatusCodes.Status401Unauthorized,
        _   => StatusCodes.Status500InternalServerError,
    };
    await ctx.Response.WriteAsJsonAsync(new { error = ex?.Message });
}));
```

---

## Phase C — VBA rewrite (`DropboxService.bas`)

The public function signatures **do not change**.  Only the internals change.
For every function being rewritten, follow the project's rollback convention:
comment-out the existing body under a `' LEGACY (pre-Bridge)` block, then write
the new body directly below it.

### C.1 — What to remove (comment out as LEGACY blocks)

| Section | Items to LEGACY-block |
|---|---|
| Constants | `APPSECRET_PLACEHOLDER`, `AUTH_URL_BASE`, `TOKEN_URL`, `ACCOUNT_URL`, `LISTENER_PORT`, `LISTENER_TIMEOUT_S`, `USE_LOCAL_LISTENER` |
| Module-level state | `m_AppKey`, `m_AppSecret`, `m_RedirectUri`, `m_NamespaceId`, `m_TeamRootPath`, `m_OAuthState`, token cache vars (`m_AccessToken` etc.) |
| Win32 declares | All `CryptProtectData`, `CryptUnprotectData`, `CopyMemory` declares |
| Section 4 | `EncryptDPAPI`, `DecryptDPAPI`, `BytesToBase64`, `Base64ToBytes` (DPAPI helpers) |
| Section 8 | `EnsureTblDropboxTokensSchema` (schema migration) |
| Section 9 | `GenerateOAuthState`, `ValidateOAuthState`, `ClearOAuthState` |
| Section 10 | `SaveTokens`, `LoadTokens`, `ClearTokens`, `IsTokenLoaded`, `TokenIsExpiring`, `GetCurrentAccessToken`, `GetCurrentRefreshToken` (all in Section 10 — Section 11 is the Phase 3b Pass-1 smoke test, leave it or drop with the other smoke tests) |
| Section 12 | `HttpRequest` (replaced by `BridgeRequest`). **NOTE: this is Section 12 "HTTP TRANSPORT", not Section 13. Section 13 is "JSON + URL HELPERS" (`ExtractJsonString`, `ExtractJsonLong`) — KEEP that one.** |
| Section 14 | `WriteListenerScript`, `PollForFile`, `WaitMilliseconds`, `EnsureOAuthTempDir`, `OpenBrowser`, `AwaitListenerRedirect`, `AwaitPasteRedirect` |
| Section 15 | `AuthenticateUser`, `ExchangeCodeForToken`, `GetCurrentAccountEmail` |
| Section 16 | `IsAccountRevoked` (REVOCATION CHECK) |
| Section 17 | `RefreshAccessToken`, `EnsureValidToken` (TOKEN REFRESH + LIFECYCLE) |

**Keep everything else unchanged:** `ALLOW_DROPBOX_WRITES`, `GuardWritesEnabled`,
`LogLocal`, `LogAuditEvent`, `EnsureConfigLoaded` / `InitializeDropboxConfig`
(simplified — see C.2), JSON helpers (`ExtractJsonString`, `ExtractJsonLong`),
path helpers (`JsonEscapePath`, `SanitizeWindowsFilename`, `DropboxBaseName`,
`BuildLocalTempPath`, `WriteBytesToFile`), `CleanupTempFiles`, `NewGuid`,
`GetFileSize`, `ReadAllBytes`, `HttpDownloadBinary` (still used for downloading
from temp links without auth — see C.4), and all smoke tests (they exercise
the public functions; the test assertions don't change).

> **After the rewrite, two kept items have no live callers:**
> `HttpDownloadBinary`'s POST/token branch (the new `OpenDocument` uses only the
> header-free GET branch — see G31) and `DropboxPathRootHeader()` (only that POST
> branch referenced it). Harmless dead code — leave a comment so a future reader
> doesn't assume they're still exercised, or remove them once the direct
> content-API path is confirmed gone. (Moot if G34 is resolved toward a bridge
> proxy, which removes the GET branch too.)

### C.2 — Simplified module-level state and `InitializeDropboxConfig`

Replace the module-level config cache with just one variable:

```vba
' Replace the 6 m_Xxx config cache vars with:
Private m_BridgeUrl    As String   ' loaded from tblDropboxConfig.BridgeUrl
Private m_ConfigLoaded As Boolean  ' unchanged

' Simplify InitializeDropboxConfig to load only BridgeUrl:
Public Sub InitializeDropboxConfig()
    Const CALLER As String = "InitializeDropboxConfig"
    If m_ConfigLoaded Then Exit Sub

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    On Error GoTo HandleError

    cn.Open PcaGetConnnectionString()
    Set rs = New ADODB.Recordset
    rs.Open "SELECT BridgeUrl FROM dbo.tblDropboxConfig WHERE ConfigID = 1", _
            cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        Err.Raise vbObjectError + 6010, CALLER, _
            "tblDropboxConfig row ConfigID=1 not found. Re-run SQL installer."
    End If
    m_BridgeUrl = Nz(rs!BridgeUrl, "")
    rs.Close
    cn.Close

    If LenB(m_BridgeUrl) = 0 Then
        Err.Raise vbObjectError + 6011, CALLER, _
            "tblDropboxConfig.BridgeUrl is empty. " & _
            "Run: UPDATE dbo.tblDropboxConfig SET BridgeUrl = N'http://<server>/api' " & _
            "WHERE ConfigID = 1;"
    End If

    m_ConfigLoaded = True
    LogLocal CALLER, "Info", "Config loaded; BridgeUrl=" & m_BridgeUrl
    Exit Sub
HandleError:
    ' ... same error surfacing pattern as the existing function
End Sub

' Accessor used by BridgeRequest
Public Function GetBridgeUrl() As String
    EnsureConfigLoaded "GetBridgeUrl"
    GetBridgeUrl = m_BridgeUrl
End Function
```

### C.3 — The new `BridgeRequest` helper (replaces `HttpRequest`)

This is the only new private function needed.  Add it in place of the LEGACY'd
Section 12 (`HttpRequest`, the old HTTP TRANSPORT section — **not** Section 13,
which holds the JSON helpers you keep):

```vba
' ============================================================================
' SECTION 12 — BRIDGE HTTP HELPER (replaces the old HTTP TRANSPORT section)
' ============================================================================
' Sends a JSON request to the TBCMSDropboxBridge service via Windows Integrated
' Auth (NTLM). The bridge owns all Dropbox credentials; VBA sends no auth
' tokens to Dropbox directly.
'
' method       : "GET" or "POST"
' endpoint     : path relative to BridgeUrl, e.g. "/metadata"
' requestBody  : JSON string (empty string for GET)
' outStatus    : HTTP status code returned
' outResponse  : response body (JSON)
' Returns True if HTTP 2xx.

Private Function BridgeRequest( _
    ByVal method As String, _
    ByVal endpoint As String, _
    ByVal requestBody As String, _
    ByRef outStatus As Long, _
    ByRef outResponse As String _
) As Boolean
    Const CALLER As String = "BridgeRequest"
    EnsureConfigLoaded CALLER

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo HttpError

    Dim url As String
    url = m_BridgeUrl & endpoint

    http.Open method, url, False   ' synchronous
    http.SetAutoLogonPolicy 0      ' AutoLogonPolicy_Always — sends NTLM automatically
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept",       "application/json"

    If LenB(requestBody) > 0 Then
        http.Send requestBody
    Else
        http.Send
    End If

    outStatus   = http.Status
    outResponse = http.ResponseText
    BridgeRequest = (outStatus >= 200 And outStatus < 300)
    Exit Function

HttpError:
    LogLocal CALLER, "Error", "WinHttp error " & Err.Number & ": " & Err.Description & _
             " calling " & method & " " & endpoint
    outStatus   = 0
    outResponse = ""
    BridgeRequest = False
End Function
```

For binary uploads (the `UploadFile` and `UploadLargeFile` case), add a second
helper:

```vba
' Sends raw bytes to the bridge upload endpoint.
' Dropbox path is passed in the X-Dropbox-Path request header.
Private Function BridgeUpload( _
    ByVal dropboxPath As String, _
    ByRef bytes() As Byte, _
    ByRef outStatus As Long, _
    ByRef outResponse As String _
) As Boolean
    Const CALLER As String = "BridgeUpload"
    EnsureConfigLoaded CALLER

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo HttpError

    http.Open "POST", m_BridgeUrl & "/file/upload", False
    http.SetAutoLogonPolicy 0
    ' Large uploads stream synchronously VBA -> bridge -> Dropbox; the default
    ' WinHttp send/receive timeout is 30s and will abort big files (G32).
    ' Args: resolve, connect, send, receive (ms). 0 = infinite resolve.
    http.SetTimeouts 0, 60000, 300000, 300000
    http.SetRequestHeader "Content-Type",  "application/octet-stream"
    http.SetRequestHeader "X-Dropbox-Path", dropboxPath
    http.Send bytes

    outStatus   = http.Status
    outResponse = http.ResponseText
    BridgeUpload = (outStatus >= 200 And outStatus < 300)
    Exit Function
HttpError:
    LogLocal CALLER, "Error", "Upload WinHttp error " & Err.Number & ": " & _
             Err.Description
    outStatus   = 0
    outResponse = ""
    BridgeUpload = False
End Function
```

### C.4 — Function-by-function rewrite table

For each function below: LEGACY-block the existing body, write the new body.

#### `OpenDocument(dropboxPath)` → ask bridge for temp link, then download from CDN

```vba
Public Function OpenDocument(ByVal dropboxPath As String) As String
    Const CALLER As String = "OpenDocument"

    ' LEGACY (pre-Bridge): called HttpDownloadBinary directly to Dropbox
    ' content API with an access token.  See the LEGACY block below.
    '
    ' Active (Bridge): ask the bridge for a 4-hour temp download link;
    ' VBA downloads from the Dropbox CDN (no auth needed for temp links).

    EnsureConfigLoaded CALLER

    Dim status As Long, resp As String
    Dim body As String
    body = "{""path"":""" & JsonEscapePath(dropboxPath) & """}"
    If Not BridgeRequest("POST", "/file/link", body, status, resp) Then
        LogLocal CALLER, "Error", "Bridge /file/link failed: HTTP " & status & _
                 " path=" & dropboxPath
        Exit Function
    End If

    Dim tempLink As String
    tempLink = ExtractJsonString(resp, "temporaryLink")
    If LenB(tempLink) = 0 Then
        LogLocal CALLER, "Error", "No temporaryLink in response: " & Left$(resp, 300)
        Exit Function
    End If

    Dim tempPath As String
    tempPath = BuildLocalTempPath(dropboxPath)

    Dim dlStatus As Long, errText As String
    Dim bytes() As Byte
    ' HttpDownloadBinary still used here — but now downloading from an
    ' unauthenticated Dropbox CDN temp link (no token required).
    If Not HttpDownloadBinary(tempLink, "", "", dlStatus, bytes, errText) Then
        LogLocal CALLER, "Error", "CDN download failed: HTTP " & dlStatus & _
                 " path=" & dropboxPath
        Exit Function
    End If

    On Error GoTo WriteError
    WriteBytesToFile tempPath, bytes
    On Error GoTo 0

    On Error Resume Next
    Shell "explorer.exe """ & tempPath & """", vbNormalFocus
    If Err.Number <> 0 Then
        LogLocal CALLER, "Error", "explorer.exe launch failed: Err=" & _
                 Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    LogLocal CALLER, "Info", "Opened " & dropboxPath & " -> " & tempPath
    OpenDocument = tempPath
    Exit Function
WriteError:
    LogLocal CALLER, "Error", "Failed writing temp file " & tempPath & _
             ": Err=" & Err.Number & " " & Err.Description
End Function
```

> **⚠ Reusing `HttpDownloadBinary` for a temp link is NOT a one-liner (corrects
> G31).** The existing `HttpDownloadBinary` (DropboxService.bas ~2050) is built
> for the Dropbox *content* API: it issues a **POST**, and unconditionally sends
> `Authorization`, `Dropbox-API-Path-Root`, and `Dropbox-API-Arg` headers. A
> `uc.dropboxusercontent.com` temp link is a plain **GET** URL that takes none of
> those headers — POSTing with Dropbox headers will not download it correctly.
> You must add a real branch (or a separate `HttpDownloadFromUrl` helper) that
> does a header-free `GET` when the token arg is empty:
> ```vba
> If LenB(bearerToken) = 0 Then
>     http.Open "GET", url, False          ' CDN temp link: no auth, no Dropbox headers
>     http.Send
> Else
>     http.Open "POST", url, False         ' content API path (unchanged)
>     http.SetRequestHeader "Authorization", "Bearer " & bearerToken
>     http.SetRequestHeader "Dropbox-API-Path-Root", DropboxPathRootHeader()
>     http.SetRequestHeader "Dropbox-API-Arg", dropboxArgJson
>     http.Send
> End If
> ```
> **⚠ Decide G34 before writing this function.** Whether `OpenDocument`
> downloads direct from the CDN (the code above) or proxies through a new bridge
> `/api/file/download` depends on whether workstations have outbound HTTPS to
> `*.dropboxusercontent.com` — which **contradicts the locked-down premise of
> this whole bridge.** Confirm workstation egress with IT first; if it's blocked,
> the direct-CDN code above (and the GET-branch surgery below) is wrong and you
> implement the proxy instead. See G34.

#### `GetMetadata(dropboxPath, outFound, outErrorDetail, outJson)` → bridge `/metadata`

```vba
Public Function GetMetadata(...) As Boolean
    ' POST /api/metadata {"path":"..."} → {"found":bool,"errorSummary":"...","rawJson":"..."}
    Dim status As Long, resp As String
    Dim body As String
    body = "{""path"":""" & JsonEscapePath(dropboxPath) & """}"
    If Not BridgeRequest("POST", "/metadata", body, status, resp) Then
        outErrorDetail = "Bridge transport failure: HTTP " & status
        GetMetadata = False
        Exit Function
    End If
    outFound       = ExtractJsonBool(resp, "found")  ' see ExtractJsonBool helper below
    outErrorDetail = ExtractJsonString(resp, "errorSummary")
    outJson        = ExtractJsonString(resp, "rawJson")
    GetMetadata    = True
End Function
```

> Implementation note: `outFound` parsing — the bridge returns `"found": true`
> as a JSON boolean.  `ExtractJsonString` won't work for booleans; either add
> an `ExtractJsonBool` helper or check if the response contains `"found":true`
> as a substring.  Recommended: add a one-line `ExtractJsonBool` helper:
> ```vba
> Private Function ExtractJsonBool(json As String, key As String) As Boolean
>     ExtractJsonBool = (InStr(json, """" & key & """:true") > 0)
> End Function
> ```

#### `ListFolder(dropboxPath)` → bridge `/folder/list`

```vba
Public Function ListFolder(ByVal dropboxPath As String) As String
    Dim status As Long, resp As String
    Dim body As String
    body = "{""path"":""" & JsonEscapePath(dropboxPath) & """}"
    BridgeRequest "POST", "/folder/list", body, status, resp
    If status >= 200 And status < 300 Then ListFolder = resp
End Function
```

#### `GetTemporaryLink(dropboxPath)` → bridge `/file/link`

```vba
Public Function GetTemporaryLink(ByVal dropboxPath As String) As String
    Dim status As Long, resp As String
    Dim body As String
    body = "{""path"":""" & JsonEscapePath(dropboxPath) & """}"
    BridgeRequest "POST", "/file/link", body, status, resp
    If status >= 200 And status < 300 Then
        GetTemporaryLink = ExtractJsonString(resp, "temporaryLink")
    End If
End Function
```

#### `UploadFile(localPath, dropboxPath, ...)` → `BridgeUpload`

```vba
Public Function UploadFile(ByVal localPath As String, _
                            ByVal dropboxPath As String, ...) As Boolean
    GuardWritesEnabled "UploadFile"
    EnsureConfigLoaded "UploadFile"

    ' LEGACY (pre-Bridge): direct HttpUploadBinary to content.dropboxapi.com
    '
    ' Active (Bridge): read bytes locally, POST to bridge /file/upload.
    ' Bridge decides single-shot vs. chunked based on file size.

    Dim bytes() As Byte
    If Not ReadAllBytes(localPath, bytes) Then
        LogAuditEvent "Upload", "Failure", caseID, documentType, dropboxPath, _
            "Failed to read source file: " & localPath
        Exit Function
    End If

    Dim status As Long, resp As String
    If Not BridgeUpload(dropboxPath, bytes, status, resp) Then
        LogAuditEvent "Upload", "Failure", caseID, documentType, dropboxPath, _
            "Bridge HTTP " & status & ": " & Left$(resp, 300)
        Exit Function
    End If

    LogAuditEvent "Upload", "Success", caseID, documentType, dropboxPath, ""
    LogLocal "UploadFile", "Info", "Uploaded " & localPath & " to " & dropboxPath
    UploadFile = True
End Function
```

#### `UploadLargeFile(...)` → same `BridgeUpload` (bridge handles chunking)

Identical to `UploadFile` above — the bridge inspects file size and routes
internally.  In the bridge architecture, VBA no longer needs a separate
`UploadLargeFile` function.  Keep the function for signature compatibility but
delegate it to `UploadFile` internally:

```vba
Public Function UploadLargeFile(...) As Boolean
    ' Delegate to UploadFile — bridge handles single vs. chunked routing.
    UploadLargeFile = UploadFile(localPath, dropboxPath, caseID, documentType)
End Function
```

#### `MoveFile(fromPath, toPath)` → bridge `/file/move`

```vba
Public Function MoveFile(ByVal fromPath As String, ByVal toPath As String, ...) As Boolean
    GuardWritesEnabled "MoveFile"
    Dim status As Long, resp As String
    Dim body As String
    body = "{""fromPath"":""" & JsonEscapePath(fromPath) & """," & _
           """toPath"":""" & JsonEscapePath(toPath) & """}"
    If Not BridgeRequest("POST", "/file/move", body, status, resp) Then
        LogAuditEvent "Move", "Failure", caseID, documentType, fromPath, "HTTP " & status
        Exit Function
    End If
    LogAuditEvent "Move", "Success", caseID, documentType, fromPath, "to " & toPath
    MoveFile = True
End Function
```

> **Forward `caseID`/`documentType` to `LogAuditEvent`.** They are part of the
> preserved (optional) signature; passing them through — as `UploadFile` does —
> keeps case/document-type attribution in `tblDropboxAuditLog`. Passing empty
> placeholders (`, ,`) silently drops that attribution for move/copy/delete.

#### `CopyFile(fromPath, toPath)` → bridge `/file/copy`

Same pattern as `MoveFile`, endpoint `/file/copy`.

> **Preserve conflict semantics.** The legacy `MoveFile`/`CopyFile` run
> `autorename=false` and deliberately "surface conflicts to the caller for
> explicit resolution" (DropboxService.bas ~2537). Collapsing the bridge result
> to a bare `True/False` loses the distinction between a transport failure and a
> Dropbox `to/conflict`. Have the bridge return the Dropbox error tag (or a
> dedicated 409 status) so the VBA caller can still tell "already exists at
> destination" apart from "the call failed", matching pre-bridge behavior.

#### `DeleteFile(dropboxPath)` → bridge `/file/delete`

```vba
Public Function DeleteFile(ByVal dropboxPath As String, ...) As Boolean
    GuardWritesEnabled "DeleteFile"
    Dim status As Long, resp As String
    Dim body As String
    body = "{""path"":""" & JsonEscapePath(dropboxPath) & """}"
    If Not BridgeRequest("POST", "/file/delete", body, status, resp) Then
        LogAuditEvent "Delete", "Failure", caseID, documentType, dropboxPath, "HTTP " & status
        Exit Function
    End If
    LogAuditEvent "Delete", "Success", caseID, documentType, dropboxPath, ""
    DeleteFile = True
End Function
```

#### `CreateFolder(dropboxPath)` → bridge `/folder/create`

```vba
Public Function CreateFolder(ByVal dropboxPath As String) As Boolean
    GuardWritesEnabled "CreateFolder"
    Dim status As Long, resp As String
    Dim body As String
    body = "{""path"":""" & JsonEscapePath(dropboxPath) & """}"
    CreateFolder = BridgeRequest("POST", "/folder/create", body, status, resp)
    If Not CreateFolder Then
        LogLocal "CreateFolder", "Error", "Bridge HTTP " & status & _
                 " path=" & dropboxPath
    End If
End Function
```

### C.5 — Simplified `StartupBootstrap`

The startup form no longer triggers OAuth.  It just pings the bridge:

```vba
Public Sub StartupBootstrap()
    Const CALLER As String = "StartupBootstrap"
    Dim errNum As Long, errDesc As String
    On Error GoTo HandleError

    InitializeDropboxConfig   ' loads BridgeUrl from SQL

    Dim status As Long, resp As String
    If Not BridgeRequest("GET", "/status", "", status, resp) Then
        MsgBox "Could not reach the Dropbox Bridge service." & vbCrLf & vbCrLf & _
               "URL: " & m_BridgeUrl & vbCrLf & _
               "HTTP status: " & status & vbCrLf & vbCrLf & _
               "Contact IT if this persists.", vbExclamation, "Dropbox Bridge Unavailable"
        LogLocal CALLER, "Error", "Bridge unreachable: HTTP " & status
        Exit Sub
    End If

    Dim bridgeStatus As String
    bridgeStatus = ExtractJsonString(resp, "status")

    If bridgeStatus = "needs_setup" Then
        MsgBox "The Dropbox Bridge service has not been configured yet." & vbCrLf & _
               "Ask IT to complete setup at: " & m_BridgeUrl & "/setup/start", _
               vbCritical, "Bridge Setup Required"
        LogLocal CALLER, "Error", "Bridge returned needs_setup"
        Exit Sub
    End If

    If bridgeStatus <> "ok" Then
        MsgBox "Dropbox Bridge returned an unexpected status: " & bridgeStatus & vbCrLf & _
               ExtractJsonString(resp, "errorDetail"), vbExclamation, "Dropbox Bridge Error"
        LogLocal CALLER, "Warn", "Bridge status=" & bridgeStatus
        Exit Sub
    End If

    LogLocal CALLER, "Info", "Bridge OK"
    Exit Sub

HandleError:
    errNum  = Err.Number
    errDesc = Err.Description
    LogLocal CALLER, "Error", "Err=" & errNum & " " & errDesc
    MsgBox "Dropbox session could not be initialized." & vbCrLf & _
           "Error: " & errDesc, vbCritical, "Startup Error"
End Sub
```

### C.6 — Update file-header status block in `DropboxService.bas`

At the very top of the module, update the `CURRENTLY IMPLEMENTED` block to add:

```
'   Bridge (Phase B):
'     - BridgeRequest / BridgeUpload (WinHttp, Windows Integrated Auth)
'     - InitializeDropboxConfig simplified to load BridgeUrl only
'     - OAuth, DPAPI, token management removed (LEGACY-blocked)
'     - StartupBootstrap simplified to ping /api/status
'     - All public function signatures unchanged
```

---

## Phase D — IIS deployment

### D.1 — Publish the service

```bash
cd dropbox-bridge
dotnet publish -c Release -r win-x64 --self-contained false -o ../publish/bridge
```

Copy the `../publish/bridge/` folder to the server, e.g.
`C:\inetpub\wwwroot\tbcms-bridge\`.

### D.2 — IIS site configuration

1. Open IIS Manager on the target server.
2. **Application Pool:** Create `TBCMSBridge`
   - .NET CLR version: **No Managed Code**
   - Identity: a **dedicated domain service account** is recommended.
     - The bridge **reads AND writes** `tblDropboxServiceToken` (token refresh
       UPSERTs it) — grant `SELECT, INSERT, UPDATE` on that table (the earlier
       "read access" note was wrong).
     - With `Integrated Security=True`, SQL sees this account's identity. If you
       keep **Network Service**, SQL sees the **machine account**
       (`DOMAIN\<bridgehost>$`) and that login must be created with the same
       rights. A dedicated account avoids this and is portable if the bridge
       moves hosts.
     - Whichever identity you pick must also have **Modify** rights on the Data
       Protection key folder (step 6).
3. **Site or Application:** Point root to `C:\inetpub\wwwroot\tbcms-bridge\`
   - Binding: HTTP on port 80, host header `tbcms-bridge.tatebywater.local`
     (or use an IP-based binding if there is no internal DNS)
4. **Authentication:** In the IIS Authentication feature for the app:
   - Enable **Windows Authentication**
   - Disable **Anonymous Authentication**
   - (Setup endpoints are localhost-only in code, not via IIS config)
5. **web.config** (auto-generated by publish, but verify):

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <system.webServer>
    <handlers>
      <add name="aspNetCore" path="*" verb="*"
           modules="AspNetCoreModuleV2" resourceType="Unspecified"/>
    </handlers>
    <security>
      <requestFiltering>
        <!-- IIS request filtering caps request size at ~28.6 MB by default;
             raise it or large uploads fail with HTTP 404.13 BEFORE reaching the
             app. This is SEPARATE from IISServerOptions.MaxRequestBodySize and
             is the in-process knob that actually gates upload size. -->
        <requestLimits maxAllowedContentLength="2147483648"/>
      </requestFiltering>
    </security>
    <aspNetCore processPath="dotnet"
                arguments=".\TBCMSDropboxBridge.dll"
                stdoutLogEnabled="true"
                stdoutLogFile=".\logs\stdout"
                hostingModel="inprocess">
      <environmentVariables>
        <environmentVariable name="ASPNETCORE_ENVIRONMENT" value="Production"/>
      </environmentVariables>
    </aspNetCore>
  </system.webServer>
</configuration>
```

6. Create `C:\ProgramData\TBCMSBridge\dpkeys\` and grant the Application Pool
   identity `Modify` rights (Data Protection key ring lives here).

7. Install `appsettings.Production.json` on the server (not in source control):

```json
{
  "Dropbox": {
    "AppSecret": "<real secret from Dropbox App Console>"
  },
  "ConnectionStrings": {
    "TateByWater": "Server=tbf-cms;Database=<prod-db>;Integrated Security=True;TrustServerCertificate=True;"
  },
  "Bridge": {
    "AllowWrites": true
  }
}
```

> This is the **production** override: it points at the production SQL host
> (`tbf-cms`, NOT the dev `awsql2022dev`) and is the **only** place
> `Bridge:AllowWrites` is `true`. A test/staging deployment keeps the dev
> connection string and `AllowWrites:false`. `tblDropboxServiceToken` and its
> one token row must exist in each target database independently (run Phase E
> setup once per environment). See G35.

### D.3 — Smoke test the IIS deployment

From the server, open a browser to:
```
http://localhost/api/status
```
Expected response (before setup): `{"status":"needs_setup","accountEmail":null,"errorDetail":null}`

---

## Phase E — One-time setup (IT admin action)

> **The entire setup round-trip must stay on `localhost`** (run a browser
> *on the IIS server itself*). Two reasons: (a) the OAuth `state` is held in a
> session cookie scoped to the host you started on — if Dropbox redirects to a
> different hostname the cookie isn't sent and the state check fails with
> "State mismatch"; (b) the callback re-checks `IsLocalRequest`. So the
> registered redirect URI is `http://localhost/api/setup/callback`, which is
> also the only `http://` (non-https) URI Dropbox will accept (see G30).

1. On the IIS server, open Edge/Chrome and navigate to:
   ```
   http://localhost/api/setup/start
   ```
   (only works from localhost — the service enforces this)

2. The browser is redirected to `https://www.dropbox.com/oauth2/authorize`.

3. Sign in with the firm's **Dropbox Business admin account** (or the dedicated
   service account).  Click **Allow**.

4. Dropbox redirects back to `http://localhost/api/setup/callback?code=...`.
   The service validates `state`, exchanges the code, fetches the account email
   (`users/get_current_account`), saves everything to `tblDropboxServiceToken`,
   and returns "Setup complete."

5. Verify: navigate to `http://localhost/api/status`.
   Expected: `{"status":"ok","accountEmail":"...", ...}`

6. From a workstation, test:
   ```
   http://tbcms-bridge.tatebywater.local/api/status
   ```
   (should return `{"status":"ok",...}` with Windows Auth passing automatically
   in Edge/Chrome on the domain).

---

## Phase F — Smoke tests

### F.1 — Bridge-side (curl or Postman from a domain workstation)

```bash
# Metadata check (path must exist in /Company)
curl -s --negotiate -u : \
  -X POST http://tbcms-bridge.tatebywater.local/api/metadata \
  -H "Content-Type: application/json" \
  -d "{\"path\":\"/Company/Clients\"}"
# Expected: {"found":true,"errorSummary":null,"rawJson":"{...}"}

# Temp link
curl -s --negotiate -u : \
  -X POST http://tbcms-bridge.tatebywater.local/api/file/link \
  -H "Content-Type: application/json" \
  -d "{\"path\":\"/Company/Clients/some-file.pdf\"}"
# Expected: {"temporaryLink":"https://uc.dropboxusercontent.com/..."}
```

### F.2 — VBA smoke tests (run in the Access Immediate window)

The existing smoke tests (`Phase3a_SmokeTest`, `Phase3c_SmokeTest`,
`Phase3d_SmokeTest`, etc.) exercise the public functions — after the VBA
rewrite they will transparently call the bridge instead of Dropbox directly.
Run them in the same order as before:

1. `? DropboxService.Phase3a_SmokeTest`
   — now tests: config loaded (BridgeUrl present), `LogLocal` writes, `LogAuditEvent` writes.
   
2. `? DropboxService.Phase3c_SmokeTest`
   — calls `GetMetadata` → bridge; verifies found/not-found tristate.
   
3. Flip `ALLOW_DROPBOX_WRITES = True`, re-import, then:
   `? DropboxService.Phase3d_SmokeTest`
   — calls `CreateFolder`, `UploadFile`, `MoveFile`, `DeleteFile` → bridge.
   Flip back to `False` after.

4. `? DocumentManagement.Phase5_E2E_HappyPathTest(30405)`
   — end-to-end with writes enabled; same test case as before.

### F.3 — Updated `Phase3b` tests

The Phase 3b tests (OAuth unit test, auth flow test) become obsolete.  Replace
`Phase3b_Pass1_SmokeTest` and `Phase3b_Pass2_AuthFlowTest` with a single bridge
connectivity test:

```vba
Public Function PhaseBridge_ConnectivityTest() As String
    On Error GoTo HandleError
    Dim stepName As String

    stepName = "1.ConfigLoad"
    InitializeDropboxConfig
    If LenB(m_BridgeUrl) = 0 Then Err.Raise vbObjectError + 6300, , "BridgeUrl empty"

    stepName = "2.StatusPing"
    Dim status As Long, resp As String
    If Not BridgeRequest("GET", "/status", "", status, resp) Then _
        Err.Raise vbObjectError + 6301, , "Bridge /status HTTP " & status

    stepName = "3.StatusValue"
    If ExtractJsonString(resp, "status") <> "ok" Then _
        Err.Raise vbObjectError + 6302, , "Bridge status=" & _
            ExtractJsonString(resp, "status")

    PhaseBridge_ConnectivityTest = "OK — bridge reachable, status=ok"
    Exit Function
HandleError:
    PhaseBridge_ConnectivityTest = "FAIL at " & stepName & ": " & _
        Err.Number & " " & Err.Description
End Function
```

---

## Phase G — Deliverable #7 (unchanged)

The `tblDropboxVerificationReport` population script (the Phase 6.5
acceptance-gate task, already defined in `dropbox-migration-plan.md`) calls
`DropboxService.GetMetadata` per document path.  After the VBA rewrite,
`GetMetadata` goes through the bridge — no changes to Deliverable #7 itself.
Implement it exactly as described in the migration plan's "Next concrete step"
section.

---

## Known gaps / decisions for the implementing session

| # | Item | Decision / action needed |
|---|---|---|
| G28 | AD group name for Windows Auth | Confirm with IT the exact `DOMAIN\GroupName` values for `AllowedAdGroups` in `appsettings.json`.  If the firm has no AD groups, allow all authenticated domain users (remove group check). |
| G29 | Bridge server hostname | Confirm the IIS server name and whether internal DNS resolves it.  Update `tblDropboxConfig.BridgeUrl` and the Dropbox App Console redirect URI accordingly. |
| G30 | `RedirectUri` for setup OAuth | **Resolved → run setup entirely on localhost.** Register `http://localhost/api/setup/callback` in the Dropbox App Console. Dropbox **rejects non-localhost `http://` redirect URIs** (they must be `https://`), and the session-cookie + loopback checks require the same host throughout — so do NOT use the `.local` hostname for the callback. See Phase E. |
| G31 | `OpenDocument` temp-link download | **Corrected — NOT a one-liner.** `HttpDownloadBinary` POSTs with `Authorization` + `Dropbox-API-Path-Root` + `Dropbox-API-Arg` headers; a CDN temp link needs a header-free **GET**. Add a real GET branch when the token arg is empty (see the note under `OpenDocument` in C.4), or adopt G34 and proxy the download through the bridge. |
| G32 | Large file upload to bridge | Files >150 MB route through `UploadFile`→`BridgeUpload`; the bridge runs Dropbox `upload_session` internally. `BridgeUpload` now sets `SetTimeouts 0, 60000, 300000, 300000` (default 30s aborts big files). **Also raised TWO body caps** — `IISServerOptions.MaxRequestBodySize` in Program.cs (the in-process knob; `ConfigureKestrel` is a no-op under in-process hosting) **and** IIS request filtering `<requestLimits maxAllowedContentLength>` in web.config (default ~28.6 MB → otherwise HTTP 404.13 before the app sees the request). Both are required for uploads >~28.6 MB. Both VBA and bridge buffer the whole file in memory today; consider streaming if multi-GB files are expected. |
| G33 | Deliverable #7 rate limiting | The verification script calls `GetMetadata` ~30,000 times — via the bridge this is one authenticated server making 30,000 Dropbox calls. Dropbox limit ≈ 1,000 req/min per app. Add a ~70 ms sleep in the (not-yet-written) Deliverable #7 loop, **and** have the bridge honor `Retry-After` on HTTP 429 (retry with backoff) rather than failing the document — neither is implemented yet. Note also that `/api/status` now calls `users/get_current_account` on **every** VBA startup (`VerifyLiveAccountEmailAsync`), so a morning login wave adds one Dropbox call per user — well under the limit, but `/status` is no longer free. |
| G34 | **Download path: direct CDN vs. proxy** | **RESOLVED → bridge proxy (2026-06-22).** Implemented `POST /api/file/download` (Program.cs / `DropboxApiClient.DownloadAsync`, content API + path-root header) and a `BridgeDownload` WinHttp helper in VBA; `OpenDocument` streams bytes through the bridge over Windows Integrated Auth and never contacts `*.dropboxusercontent.com`. This matches the locked-down premise (no per-workstation internet egress required) and made the G31 GET-branch surgery on `HttpDownloadBinary` unnecessary (that helper is now LEGACY-blocked). `GetTemporaryLink` is retained (bridge `/file/link`) for link-distribution scenarios only. **Confirm with IT** the bridge host itself has outbound HTTPS to Dropbox (it does the CDN fetch now). |
| G35 | **Per-environment DB / token isolation** | The bridge owns a real Dropbox service-account token. Test and production are separate deployments with separate `appsettings.Production.json`, separate SQL hosts (`awsql2022dev` vs `tbf-cms`), and a separately-provisioned `tblDropboxServiceToken` row each. A bridge pointed at the wrong DB violates "two environments, never mix." **Also:** the test bridge's token can write to the real production `/Company` — `Bridge:AllowWrites=false` on test is the guard (G37). |
| G36 | **HTTP vs HTTPS on the LAN** | VBA→bridge is plaintext HTTP/80: document bytes, 4-hour unauthenticated temp links, and NTLM all cross the LAN in clear. For a law firm handling confidential client documents, recommend HTTPS with an internal/AD-CS cert and update `tblDropboxConfig.BridgeUrl` to `https://`. Decision needed before production cutover. |
| G37 | **Server-side write guard + access scope** | Implemented `Bridge:AllowWrites` (`GuardWrites` in Program.cs) as the bridge-side mirror of `ALLOW_DROPBOX_WRITES` — the VBA flag alone protects nothing once the bridge exists. Open item: the bridge gates only on Windows auth, so **any** TBCMS-domain user can read/temp-link/delete **any** `/Company` path via a direct HTTP call, and Dropbox-native audit shows only the service account (SQL `tblDropboxAuditLog` retains the real user). Accept this risk or add per-path / per-operation authorization. |
| G38 | `ListFolder` pagination | Neither the bridge `ListFolderAsync` nor the legacy VBA appears to handle `has_more` / `list_folder/continue`. Folders with >~2000 entries truncate silently. Confirm no document folder exceeds this, or implement continuation in the bridge and return the concatenated result. |
| G39 | Concurrent token refresh | Multiple simultaneous VBA requests near token expiry can trigger concurrent refreshes + racing SQL writes. Serialize refresh in `DropboxTokenService` with a **`static`** `SemaphoreSlim` — the service is scoped, so a per-instance semaphore serializes nothing (see B.4). |
| G40 | Data Protection key-ring scope | `ProtectKeysWithDpapi(protectToLocalMachine: true)` is now set in Program.cs. Without it the key ring is bound to the app-pool identity's profile — re-introducing the G27 profile-binding failure server-side. Back up `C:\ProgramData\TBCMSBridge\dpkeys` and keep the app-pool identity stable; losing the ring means re-running Phase E setup. |

---

## Commit/branch guidance

- Branch from `main` (the prior `claude/ms-access-dropbox-auth-udqa7w` branch was
  merged via PR #1 and should not be reused)
- Use commit prefixes matching the project convention:
  - `feat: Phase B` — new bridge service files
  - `feat: Phase C` — VBA rewrite commits
  - `docs: Phase B` — this plan + SQL installer additions
- Commit the SQL installer changes (`Section 9`) before the VBA changes so the
  schema is ready before the VBA module tries to load `BridgeUrl`
- Do **not** push Phase C VBA changes until the bridge is deployed and
  Phase F.1 bridge smoke tests pass on the actual server
- `ALLOW_DROPBOX_WRITES` remains `False` in committed source throughout

---

## Acceptance criteria (before production cutover)

- [ ] `GET /api/status` from a domain workstation returns `{"status":"ok"}`
- [ ] `PhaseBridge_ConnectivityTest` returns `"OK"` in the Immediate window
- [ ] `Phase3c_SmokeTest` passes (GetMetadata through bridge)
- [ ] `Phase3d_SmokeTest` passes (all write ops through bridge)
- [ ] `Phase5_E2E_HappyPathTest(30405)` passes end-to-end
- [ ] A new staff member (no prior token) opens Access → no OAuth prompt → Dropbox features work immediately
- [ ] `ALLOW_DROPBOX_WRITES = False` in committed source
- [ ] `Bridge:AllowWrites = false` in committed `appsettings.json`, and a direct `POST /api/file/delete` against the **test** bridge returns **403** (server-side guard, G37)
- [ ] Test and production bridges point at different SQL hosts, each with its own `tblDropboxServiceToken` row (G35)
- [ ] HTTP-vs-HTTPS decision made for VBA→bridge traffic before cutover (G36)
- [ ] Download path decision made (direct CDN vs. bridge proxy) and workstation egress to Dropbox confirmed with IT if staying direct (G34)
- [ ] `tblDropboxTokens` table is no longer written to (token storage moved to `tblDropboxServiceToken` on the server)
