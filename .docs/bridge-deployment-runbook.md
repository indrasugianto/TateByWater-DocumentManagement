# TBCMS Dropbox Bridge — Production Deployment Runbook (TBF-CMS)

> Target: install the `TBCMSDropboxBridge` service on **TBF-CMS** (the same
> Windows server that hosts the production SQL instance) under **IIS + HTTPS**,
> with **`Bridge:AllowWrites = true`**.
>
> This deploys the bridge **infrastructure only**. It is **non-disruptive** to
> the live application: production TBCMS keeps running on `S:\` until the
> separate frontend cutover (migration plan Phase 7). Deploying the bridge does
> not migrate any data and does not change how the current `.accde` behaves.
>
> Companion docs: `.docs/dropbox-bridge-plan.md` (design), `.docs/dropbox-
> migration-plan.md` (overall migration / Phase 7 cutover).

---

## 0. What the bridge actually needs on TBF-CMS

- **.NET 8 ASP.NET Core Hosting Bundle** (provides the runtime + the IIS
  AspNetCoreModuleV2). Framework-dependent publish — no runtime is bundled.
- **One SQL object** in the production database: `dbo.tblDropboxServiceToken`.
  The bridge reads the Dropbox AppKey/AppSecret/namespace from
  `appsettings.Production.json` — **not** from SQL — so it does **not** need
  `tblDropboxConfig`, the audit tables, or the stored procedures. (Those belong
  to the frontend cutover, Phase 7.)
- **A server-side OAuth token** provisioned once (Phase E), stored encrypted in
  `tblDropboxServiceToken`.

> **Domain note:** the dev workstation cannot reach TBF-CMS over Windows
> integrated auth (different/untrusted domain), so every SQL/IIS step below is
> run **on TBF-CMS** (RDP) or via SSMS/IIS Manager with a login that has rights
> there. The bridge itself works because, running on TBF-CMS, its app-pool
> identity authenticates to the **local** SQL instance.

> **⚠ Security (you chose `AllowWrites=true`):** once the token is provisioned
> and writes are enabled, the bridge can modify the entire real `/Company`
> tree, and it gates only on Windows auth — so **any** authenticated TBCMS
> domain user who can reach the site can upload/move/delete via a direct HTTP
> call (plan G37). Mitigate by (a) restricting the IIS site to a specific AD
> group (G28), and (b) protecting the token + Data-Protection key ring (steps 5,
> 8). Also rotate the Dropbox AppSecret + SQL user before go-live (CLAUDE.md).

---

## 1. Prerequisites (on TBF-CMS)

1. **IIS + Windows Authentication role feature** installed (Server Manager →
   Add Roles/Features → Web Server (IIS) → Security → Windows Authentication).
2. **.NET 8 Hosting Bundle**: download `dotnet-hosting-8.0.x-win.exe` from
   https://dotnet.microsoft.com/download/dotnet/8.0 → install → then:
   ```
   net stop was /y
   net start w3svc
   ```
   Verify: `dotnet --list-runtimes` shows `Microsoft.AspNetCore.App 8.0.x`.

## 2. Copy the published build

Copy the **contents** of the dev box's `publish\bridge\` folder to, e.g.:
```
C:\inetpub\wwwroot\tbcms-bridge\
```
(`web.config` is included and already sets `hostingModel="inprocess"`,
`ASPNETCORE_ENVIRONMENT=Production`, and the 2 GB upload cap.)

## 3. Create the token table in the production DB

In SSMS on TBF-CMS, first confirm the DB name, then create the one table
(idempotent — safe to re-run, touches no existing data):

```sql
SELECT name FROM sys.databases WHERE name LIKE 'TateByWater%';   -- note the prod DB name

USE [<prod-db-name>];
GO
IF OBJECT_ID('dbo.tblDropboxServiceToken') IS NULL
CREATE TABLE dbo.tblDropboxServiceToken (
    TokenID      int           NOT NULL PRIMARY KEY,   -- always 1 (singleton)
    AccessToken  nvarchar(max) NOT NULL,               -- Data-Protection-encrypted
    RefreshToken nvarchar(max) NOT NULL,               -- Data-Protection-encrypted
    ExpiresAtUtc datetime2     NOT NULL,
    AccountEmail nvarchar(200) NULL,
    UpdatedAtUtc datetime2     NOT NULL
        CONSTRAINT DF_DropboxServiceToken_UpdatedAtUtc DEFAULT (SYSUTCDATETIME()),
    SetupByUser  nvarchar(200) NULL,
    CONSTRAINT CK_DropboxServiceToken_SingleRow CHECK (TokenID = 1)
);
GO
```

> Do **not** run the full `Dropbox-Migration-SQL-Install.sql` against production —
> it is destructive on re-run and belongs to the Phase 7 cutover.

## 4. App pool + SQL grant

1. IIS Manager → Application Pools → **Add**: name `TBCMSBridge`,
   **.NET CLR version = No Managed Code**, Start mode Automatic.
2. **Identity** (Advanced Settings → Process Model → Identity). Two options:
   - **Dedicated domain service account** (recommended, portable):
     e.g. `TATEBYWATER\svcTBCMSBridge`.
   - **ApplicationPoolIdentity** — works against the **local** SQL instance via
     the virtual account `IIS APPPOOL\TBCMSBridge`.
3. Grant that identity rights on the prod DB (SSMS, run on TBF-CMS). Example for
   the local app-pool identity:
   ```sql
   USE [master];
   IF SUSER_ID(N'IIS APPPOOL\TBCMSBridge') IS NULL
       CREATE LOGIN [IIS APPPOOL\TBCMSBridge] FROM WINDOWS;
   GO
   USE [<prod-db-name>];
   IF USER_ID(N'IIS APPPOOL\TBCMSBridge') IS NULL
       CREATE USER [IIS APPPOOL\TBCMSBridge] FOR LOGIN [IIS APPPOOL\TBCMSBridge];
   GRANT SELECT, INSERT, UPDATE ON dbo.tblDropboxServiceToken TO [IIS APPPOOL\TBCMSBridge];
   GO
   ```
   (Swap in `TATEBYWATER\svcTBCMSBridge` if using the domain account.)

## 5. Data Protection key ring

The bridge encrypts the stored token with **machine-scope** DPAPI and persists
the key ring to disk. Create the folder and grant the app-pool identity Modify:
```
mkdir C:\ProgramData\TBCMSBridge\dpkeys
icacls C:\ProgramData\TBCMSBridge\dpkeys /grant "IIS APPPOOL\TBCMSBridge:(OI)(CI)M"
```
> Back this folder up. Losing it (or changing the app-pool identity such that it
> can't decrypt) requires re-running the Phase E setup (plan G40).

## 6. `appsettings.Production.json` (create on the server — never commit)

Create `C:\inetpub\wwwroot\tbcms-bridge\appsettings.Production.json`:
```json
{
  "Dropbox": {
    "AppSecret": "<real secret from Dropbox App Console>"
  },
  "ConnectionStrings": {
    "TateByWater": "Server=tbf-cms;Database=<prod-db-name>;Integrated Security=True;TrustServerCertificate=True;"
  },
  "Bridge": {
    "AllowWrites": true
  }
}
```
- **AppSecret**: copy from the Dropbox App Console (Settings → *App secret* →
  Show) for app key `dqleswbnux8k3m5`. Alternatively set the environment
  variable `Dropbox__AppSecret` on the app pool instead of putting it in the
  file. **Do not commit this value anywhere.**
- `NamespaceId` (`14334595683`) and `RedirectUri`
  (`http://localhost/api/setup/callback`) come from the base `appsettings.json`
  and are correct as-is — confirm the namespace matches the client tenant.
- The base file's connection string points at the dev host; this Production
  override **must** point at TBF-CMS (plan G35 — never mix environments).

## 7. HTTPS certificate + IIS site

1. Obtain a server cert for the bridge hostname (internal AD CS preferred;
   self-signed acceptable if workstations trust it). Decide the hostname, e.g.
   `tbcms-bridge.tatebywater.local`, and add an internal **DNS A record** →
   TBF-CMS. Import the cert into **Local Machine → Personal**.
2. IIS Manager → **Add Website** (or Application): name `TBCMSBridge`, physical
   path `C:\inetpub\wwwroot\tbcms-bridge\`, app pool `TBCMSBridge`.
3. **Bindings:**
   - `https` : 443, host `tbcms-bridge.tatebywater.local`, select the cert. *(operational traffic)*
   - `http`  : 80, host **blank** (or `localhost`) — **temporary**, for the
     one-time localhost OAuth setup in step 9. Remove it afterward.
4. **Authentication** (site → Authentication): **Enable Windows
   Authentication**, **Disable Anonymous Authentication**.
5. (Optional, recommended given AllowWrites=true) restrict the site to a
   specific AD group via *Authorization Rules* or URL Authorization (G28).

## 8. Smoke-test the install (before OAuth)

On the server: browse to `http://localhost/api/status`.
Expected (no token yet): `{"status":"needs_setup","accountEmail":null,"errorDetail":null}`
- 500/HTTP errors → check `C:\inetpub\wwwroot\tbcms-bridge\logs\stdout*` (stdout
  logging is on in web.config) and Event Viewer → Application.

## 9. Phase E — one-time OAuth setup (provisions the token)

1. In the **Dropbox App Console** (app key `dqleswbnux8k3m5` → Settings →
   OAuth 2 → Redirect URIs), add exactly:
   ```
   http://localhost/api/setup/callback
   ```
2. **On TBF-CMS** (RDP), open a browser to:
   ```
   http://localhost/api/setup/start
   ```
   Sign in with the firm's **Dropbox Business admin** account → **Allow**. You
   should see "Setup complete."
3. Verify `http://localhost/api/status` → `{"status":"ok","accountEmail":"…"}`.
   The encrypted token is now in `tblDropboxServiceToken` on the prod DB.
4. **Remove the temporary `http` binding** from step 7.3 (leave only `https`).

## 10. Verify over HTTPS from a workstation

From a domain workstation (Windows auth flows automatically in Edge/Chrome):
```
https://tbcms-bridge.tatebywater.local/api/status      → {"status":"ok",...}
```
`curl` example:
```
curl -s --negotiate -u : https://tbcms-bridge.tatebywater.local/api/status
curl -s --negotiate -u : -X POST https://tbcms-bridge.tatebywater.local/api/metadata ^
  -H "Content-Type: application/json" -d "{\"path\":\"/Company\"}"
```

---

## 11. NOT part of this deployment — frontend cutover (Phase 7)

The production `.accde` still uses `S:\` and will not call the bridge until the
separate cutover. When you do cut over:
- Run the full `Dropbox-Migration-SQL-Install.sql` (Sections 1–8) against the
  prod DB: Phase 2 schema, the `S:\`→`/Company/` path migration, and the
  stored-procedure rewrites. This is the destructive/data-changing step.
- Add `BridgeUrl` to `tblDropboxConfig` (installer Section 9) and set it to
  `https://tbcms-bridge.tatebywater.local/api`.
- Rebuild the production `.accde` from the source DB with the bridge `.bas`
  modules imported and `ALLOW_DROPBOX_WRITES = True`.
- Then the staff frontend will route document operations through this bridge.

## 12. Rollback (bridge)

- Stop/disable the IIS site + app pool. The `S:\` frontend is unaffected.
- The `tblDropboxServiceToken` table and `dpkeys` folder can be left in place or
  dropped/deleted. To fully revert the VBA later, flip
  `#Const PREBRIDGE_LEGACY = True` in `DropboxService.bas` and rebuild.
