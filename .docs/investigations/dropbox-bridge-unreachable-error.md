# Investigation: "Could not reach the Dropbox Bridge service" on `frmHome`

**Scope note:** flat files (`.docs/reconciliation-forms-null-calculation-investigation.md`) are this repo's existing convention for this kind of note — there is no pre-existing `.docs/investigations/` or `.docs/adr/` folder. The task specified `.docs/investigations/`, so that's where this lives; flagging the mismatch rather than silently picking one.

This is a read-only investigation. No source file was modified, no SQL was run against any database, and no `.accde`/`.accdb` was opened. Two local, non-destructive network probes were run from this machine (a DNS lookup and a TCP connect attempt) — both cited with output below.

---

## Verdict, up front

**No code defect.** `StartupBootstrap` and `BridgeRequest` in `DropboxService.bas` behaved exactly as designed — this dialog *is* the module's designed failure surface for "the bridge didn't answer," not a bug in the handler.

**Proximate cause: configuration.** Whatever SQL Server database this session's frontend is linked to has `tblDropboxConfig.BridgeUrl` set to `http://tbcms-bridge.tatebywater.local/api` — the literal seed placeholder the installer writes — and nothing answers there.

**Standing cause: infrastructure never provisioned (confirmed empirically this session).** `tbcms-bridge.tatebywater.local` does not resolve on this network (NXDOMAIN from the actual corporate DNS server — see §5). This matches the repo's own last-recorded status: the bridge's IIS deployment (Phase D) and OAuth setup (Phase E) are logged as **not done**, and the DNS A record for this hostname is a step in the *production* deployment runbook that has not been executed.

Both causes are real and connected: the config points at a name that (as far as this repo's history shows) was never stood up. Whether *this specific* database's `BridgeUrl` was always the un-updated placeholder, or reverted to the placeholder very recently via a destructive installer re-run, is the one open question — see §6 and §8 for the discriminating check.

---

## 1. Exact source of the dialog text

`Dropbox-Migration/DropboxService.bas`, function `StartupBootstrap`, bridge-era branch (`#Else` side of `#If PREBRIDGE_LEGACY`):

```
DropboxService.bas:4045-4052
    stepName = "BridgeStatus"
    Dim status As Long, resp As String
    If Not BridgeRequest("GET", "/status", "", status, resp) Then
        LogLocal CALLER, "Error", "Bridge unreachable: HTTP " & status
        StartupBootstrap = "Could not reach the Dropbox Bridge service." & vbCrLf & _
            "URL: " & m_BridgeUrl & vbCrLf & _
            "HTTP status: " & status & vbCrLf & _
            "Contact IT if this persists."
        Exit Function
    End If
```

This produces the string body the user saw almost verbatim: `"Could not reach the Dropbox Bridge service." / "URL: <BridgeUrl>" / "HTTP status: <n>" / "Contact IT if this persists."` The reported dialog capitalizes "HTTP Status" and inserts a blank line before "Contact IT" that the code above does not — cosmetic only (see §7), not a sign of a different source.

**The dialog's title ("Dropbox session"), its leading "⚠", and its closing sentence — "TBCMS will continue to load, but document-open operations will not work until this is resolved." — are NOT produced by this string.** That wrapper text does not appear anywhere in `DropboxService.bas`, `DocumentManagement.bas`, the SQL installer, `dropbox-bridge/`, or any tracked or untracked `.docs/*.md` file. See §9 for where it must actually live.

`StartupBootstrap` itself is called from `frmHome`'s `Form_Open` event (per the module's own header comment, `DropboxService.bas:3926`: *"Form_Open in the Access startup form (frmHome) calls StartupBootstrap..."*), and this is the module's documented startup hook — not something invoked ad hoc elsewhere. Checked negative: `DocumentManagement.bas` was grepped for "bridge"/"StartupBootstrap"/this dialog text and has none — it never calls `StartupBootstrap` or constructs any part of this message; the bootstrap/dialog logic lives exclusively in `DropboxService.bas`.

---

## 2. What "HTTP Status: 0" means here

`status` in the snippet above is the `outStatus` parameter of `BridgeRequest`, the private helper `StartupBootstrap` calls. `BridgeRequest` (`DropboxService.bas:1565-1618`) opens a synchronous `WinHttp.WinHttpRequest.5.1` request (`CreateObject("WinHttp.WinHttpRequest.5.1")`, line 1585) to `m_BridgeUrl & endpoint` and sends it. Its error handler is the tell:

```
DropboxService.bas:1612-1618
HttpError:
    LogLocal CALLER, "Error", "WinHttp error " & Err.Number & ": " & Err.Description & _
             " calling " & method & " " & endpoint
    outStatus = 0
    outResponse = ""
    BridgeRequest = False
```

**`outStatus` is hard-set to `0` only when the `On Error GoTo HttpError` handler fires** — i.e., when the WinHttp call itself raised a VBA runtime error before any HTTP response was received. This is not a status code the bridge (or Dropbox) returned; it's a VBA-side sentinel meaning "no response arrived at all." Concretely, WinHttp raises this way for DNS resolution failure, connection refused, connection timeout, or TLS handshake failure — exactly the note already in this repo's task brief. A real HTTP error from a reachable server (401, 404, 500, 503) would instead show as that literal number, because `outStatus = http.Status` runs and populates a real code before the function returns normally (line 1597). **"HTTP Status: 0" is proof the request never reached a server that answered — not proof of what specifically blocked it (DNS vs. TCP refused vs. timeout vs. TLS).**

By design, that finer detail is not thrown away — it just doesn't go into the dialog. The same `HttpError` handler logs the real WinHttp error code and description before zeroing the status:

```
DropboxService.bas:1612-1614
HttpError:
    LogLocal CALLER, "Error", "WinHttp error " & Err.Number & ": " & Err.Description & _
             " calling " & method & " " & endpoint
```

That distinguishes DNS failure, connection-refused, and timeout from each other (different WinHttp error codes/descriptions) in a way "HTTP Status: 0" alone cannot. See §8 step 1 for where this lands and how to read it.

---

## 3. Architecture: CLAUDE.md is stale here — this repo now routes through a bridge, not Dropbox directly

CLAUDE.md's "VBA never builds a path" / OAuth-and-DPAPI description is **the pre-bridge design**, now the *rollback* path. The active code path is a rewrite that inserts a local ASP.NET Core service, `TBCMSDropboxBridge`, between VBA and Dropbox:

```
DropboxService.bas:8-31 (module header, "BRIDGE REWRITE")
  "...go through the internal TBCMSDropboxBridge REST service over Windows
   [Integrated Auth]... The bridge owns one server-side service-account token."
DropboxService.bas:107-115
  ' False  = bridge era (current). VBA proxies every Dropbox call through the
  '          TBCMSDropboxBridge service; all OAuth/DPAPI/token blocks below are
  '          compiled OUT...
  #Const PREBRIDGE_LEGACY = False
```

`PREBRIDGE_LEGACY = False` is the value **currently committed on `main`** — the bridge path is the live, default-compiled behavior, not a future plan. VBA's only remaining Dropbox-related config is the bridge's own URL:

```
DropboxService.bas:461-497 — InitializeDropboxConfig (bridge era)
    rs.Open "SELECT BridgeUrl FROM dbo.tblDropboxConfig WHERE ConfigID = 1", ...
    m_BridgeUrl = Nz(rs!BridgeUrl, "")
    If LenB(m_BridgeUrl) = 0 Then
        Err.Raise vbObjectError + 6089, , "tblDropboxConfig.BridgeUrl is empty. Run: " & _
            "UPDATE dbo.tblDropboxConfig SET BridgeUrl = N'http://<server>/api' ..."
```

Answering the task's framing directly: this is **not** (a) an undocumented-elsewhere component — it *is* documented, extensively, in `.docs/dropbox-bridge-plan.md` and `.docs/bridge-deployment-runbook.md`; it is **not** (c) a stale leftover — it's the current live design. It is closest to **(b)**, with a twist: the *bridge service itself* is real, in-repo, and wired into the active VBA — but the specific server/DNS name the error names (`tbcms-bridge.tatebywater.local`) is an environment/infrastructure artifact that (per this repo's own records) has never actually been deployed. The bridge source is in-repo (`dropbox-bridge/`, tracked: `Program.cs`, `Services/DropboxApiClient.cs`, `Services/DropboxTokenService.cs`, `Models/*.cs`, `TBCMSDropboxBridge.csproj`, `appsettings*.json`, `web.config`), but nothing in the repo shows an actual running instance behind that hostname.

The bridge's own architecture diagram confirms VBA talks only to the bridge, and the bridge alone talks to Dropbox:

```
.docs/dropbox-bridge-plan.md:77-89
MS Access VBA (DropboxService.bas)
        │  WinHttp, Windows Integrated Auth (NTLM)
        ▼
TBCMSDropboxBridge   (ASP.NET Core 8 Minimal API, hosted on IIS)
        │  HTTPS, Bearer token, Dropbox-API-Path-Root header
        ▼
Dropbox Business API  (api.dropboxapi.com / content.dropboxapi.com)
```

---

## 4. Where `tbcms-bridge.tatebywater.local` comes from

It is a **placeholder string written by the SQL installer**, not a value anyone typed into VBA:

```
Dropbox-Migration-SQL-Install.sql:2343-2362  (Section 9.1)
if not exists (... column 'BridgeUrl' on tblDropboxConfig ...)
begin
    alter table dbo.tblDropboxConfig add BridgeUrl nvarchar(500) null;
    exec ('update dbo.tblDropboxConfig
           set    BridgeUrl = N''http://tbcms-bridge.tatebywater.local/api''
           where  ConfigID = 1;');
    print N'    SECTION 9.1: BridgeUrl column added and seeded (placeholder URL).';
end;
```

The installer's own closing banner tells IT to change it:

```
Dropbox-Migration-SQL-Install.sql:2404-2405
print N'  Don''t forget: UPDATE dbo.tblDropboxConfig SET AppSecret = N''<real>'' WHERE ConfigID = 1;';
print N'  Bridge: UPDATE dbo.tblDropboxConfig SET BridgeUrl = N''http://<server>/api'' WHERE ConfigID = 1;';
```

This same hostname reappears once more, in the **production** deployment runbook, as the *planned* real bridge hostname once actually deployed:

```
.docs/bridge-deployment-runbook.md:157-160 (Section 7, step 1)
   Obtain a server cert for the bridge hostname ... Decide the hostname, e.g.
   `tbcms-bridge.tatebywater.local`, and add an internal DNS A record → TBF-CMS.
```

But that same runbook is explicit that this deployment has **not** happened as part of the current (test-phase) work:

```
.docs/bridge-deployment-runbook.md:211-214 (Section 11)
## 11. NOT part of this deployment — frontend cutover (Phase 7)
The production `.accde` still uses `S:\` and will not call the bridge until the
separate cutover.
```

And the bridge plan's own last-recorded status snapshot says the deploy step was never done:

```
.docs/dropbox-bridge-plan.md:38-42 (▶ Implementation status, 2026-06-22)
Not done (deployment / hardware-dependent): Phases D (IIS deploy) and
E (one-time OAuth setup) are manual server actions...
```

**Net: the exact URL in the error dialog is the installer's un-replaced default, pointing at a hostname whose only documented "deployment" step (DNS A record + IIS site) has, per every status record in this repo, never been executed.**

---

## 5. Empirical confirmation from this machine (this session)

Two read-only network probes, run from this machine, whose resolver is the firm's own corporate DNS server (`TBF-SRVR19.TATEBYWATER.COM`) — i.e. this probe ran on the firm's network, using the firm's DNS, not from an unrelated/off-network host:

```
$ nslookup tbcms-bridge.tatebywater.local
*** TBF-SRVR19.TATEBYWATER.COM can't find tbcms-bridge.tatebywater.local: Non-existent domain
Server:  TBF-SRVR19.TATEBYWATER.COM
Address:  10.0.0.5
```

`tbcms-bridge.tatebywater.local` is **NXDOMAIN on the firm's own DNS, today** — direct, primary evidence (not just "no mention in the docs") that no A record for this exact name currently resolves on this network. Worth being precise about what this does and doesn't rule out: **the corporate resolver answered** (it returned a definitive NXDOMAIN, not a timeout) — so this is not a VPN/off-network/resolver-unreachable problem. The name specifically is simply not registered.

There's also a smaller, easy-to-miss detail worth naming: the firm's real DNS domain, per the resolver's own name, is `TATEBYWATER.COM` — but the hostname in question is `tbcms-bridge.tatebywater.**local**` (a different suffix). The deployment runbook itself only ever offers this as an example, never a fixed decision: *"Decide the hostname, **e.g.** `tbcms-bridge.tatebywater.local`, and add an internal DNS A record"* (`.docs/bridge-deployment-runbook.md:158-159`, emphasis added). So the installer's seed value is best read as an illustrative placeholder in a suffix that may not even match the firm's real internal DNS convention, not a specific target that was promised and failed to materialize. This corroborates §4's document-based finding (Phase D never done) rather than merely repeating it, and adds the "which suffix" nuance the docs alone don't settle.

```
$ curl -s -m 3 -o - -w "\nHTTP_CODE:%{http_code}\n" http://localhost:8088/api/status
HTTP_CODE:000
$ netstat -ano | grep 8088
(no output)
```

Nothing is listening on `localhost:8088` either — relevant because the *dev-environment* `BridgeUrl` value recorded on `awsql2022dev` as recently as 2026-08-27 was `http://localhost:8088/api` (see §6), matching the bridge's local dev launch profile (`dropbox-bridge/Properties/launchSettings.json`, `"applicationUrl": "http://localhost:8088"`). Per the (unmerged, see §9) frontend update guide, the dev bridge is started by hand with `dotnet run` — it is not a Windows Service and does not survive a reboot. **So even the "correct," most-recently-verified dev `BridgeUrl` would currently also fail with HTTP Status 0**, for a different, purely local reason (nobody has the bridge process running right now).

---

## 6. Config timeline — why the placeholder specifically, and the one open question

A separate research note, produced one day before this investigation and **not yet merged to `main`** (branch `docs/tbcms-frontend-dropbox-update-guide`, commit `54e5bc3`), recorded a **live** `sqlcmd` check against `awsql2022dev/TateByWater` on 2026-08-27:

| Check | Result (2026-08-27) |
|---|---|
| `tblDropboxConfig.BridgeUrl` column exists | Yes |
| `tblDropboxConfig.BridgeUrl` current value | `http://localhost:8088/api` |
| `tblDropboxServiceToken` account | `sugianto@tatebywater.com`, last updated 2026-06-29 |

That is **not** the placeholder — as of yesterday, the dev DB's `BridgeUrl` had already been pointed at a local dev bridge instance. The user's dialog today shows the placeholder instead. Two things changed between those two points that are worth naming precisely:

1. **A same-morning commit touched the installer that reseeds this exact column.** `2da1417` ("feat: back up tblCaseDocuments paths before SECTION 5 rewrite"), committed **2026-08-28 09:14** — today, by this repo's own user — edits `Dropbox-Migration-SQL-Install.sql`. Section 1 of that installer unconditionally `DROP`/`CREATE`s `tblDropboxConfig` on every run (installer header, line 55: *"DESTRUCTIVE on re-run"*); Section 9.1 (§4 above) only re-adds and re-seeds `BridgeUrl` to the placeholder **when the column is found missing** — i.e., exactly the state right after a Section 1 `DROP`/`CREATE`. This is a plausible, dated mechanism for the placeholder to have reappeared, **if** the full installer (not just the script file) was re-run against whatever database backs this session's frontend, this morning.
2. Alternatively, this session's frontend may simply be linked to a **different** database than the one checked on 2026-08-27 (a different `.accdb`/`.accde`, a fresh restore, or — contrary to CLAUDE.md's "no script runs against production until cutover" rule — production `tbf-cms` with Section 9 having been applied there too). Nothing in the repo confirms or rules this out.

**This could not be resolved without running a live query, which this investigation's instructions prohibit.** The single discriminating check, to hand to IT/dev:

```sql
SELECT BridgeUrl, ConfigID FROM dbo.tblDropboxConfig WHERE ConfigID = 1;
-- run against whichever SQL Server + database the affected .accde/.accdb is actually linked to
```

- If it returns the placeholder `http://tbcms-bridge.tatebywater.local/api` → confirms a reseed happened (installer re-run) or it was never customized on this DB; the fix is a one-line `UPDATE` (§8).
- If it returns `http://localhost:8088/api` (or something else non-placeholder) → the placeholder in the error dialog came from a *different* database than the one just checked, meaning the environment question ("which `.accde`, which SQL Server, which database") is still open and needs to be answered directly, not inferred.

One thing this rules out on its own: because the dialog is specifically "bridge unreachable" (`StartupBootstrap`'s `BridgeStatus` step) rather than "Dropbox session could not be initialized at step 'InitializeDropboxConfig'" (the earlier step, which raises if `BridgeUrl` comes back **empty** — `DropboxService.bas:489-492`), the config *read* succeeded and returned a non-empty URL. So whatever database is in play does have the Section 9 schema applied and a populated `BridgeUrl` — it's the *value*, not the presence, that's in question.

---

## 7. Minor: exact wording differences (not evidence of a different source)

The reported dialog reads `"HTTP Status: 0"` (capital S, and apparently a blank line before "Contact IT"); the code at `DropboxService.bas:4051-4052` emits `"HTTP status: " & status` (lowercase s) directly followed by `"Contact IT if this persists."` with no blank line in between. This is cosmetic — most likely the reporter's own retyping of the dialog, or reformatting inside whatever wrapper produced the surrounding "Dropbox session" box (§9) — and should not be read as evidence of a second, different error-construction site. No other string in the repo produces this phrase with different wording, so `StartupBootstrap` (§1) remains the sole identified source of the message body.

---

## 8. Ranked next steps

1. **Read the latest Error row in the frontend's own local `tblDropboxLog` — cheapest check, needs no SQL Server access at all.** `tblDropboxLog` is a per-frontend local DAO table, self-created by `EnsureLocalLogTable` inside the same Access file the user hit the dialog in (not on SQL Server). Two rows in it answer both open questions in this note directly:
   - The `InitializeDropboxConfig` Info row logs the exact URL that session loaded — `"Config loaded; BridgeUrl=" & m_BridgeUrl` (`DropboxService.bas:497`) — settling §6's "which database, which value" question without touching SQL Server.
   - The `BridgeRequest` Error row logs the real WinHttp error number and description before it gets zeroed to status 0 — `"WinHttp error " & Err.Number & ": " & Err.Description & " calling " & method & " " & endpoint` (`DropboxService.bas:1613`) — which distinguishes DNS failure from connection-refused from timeout, the ambiguity §2 notes "HTTP Status: 0" alone can't resolve. The migration plan already prescribes exactly this recovery step: *"To recover the real error from any future startup failure, read the latest `tblDropboxLog` Error row (it captures `Err=<n> <desc>` before any clear)"* (`.docs/dropbox-migration-plan.md:40`).
2. **If SQL Server access is available anyway, run both of these together** (same read-only round trip, against whichever database the affected `.accde`/`.accdb` is actually linked to) — the second corroborates which branch of §6 is true:
   ```sql
   SELECT BridgeUrl, ConfigID FROM dbo.tblDropboxConfig WHERE ConfigID = 1;
   SELECT COUNT(*) FROM dbo.tblDropboxAuditLog;
   ```
   The audit-log row count was 33 on 2026-08-27 (§6). Section 1 of the installer unconditionally `DROP`/`CREATE`s that table on every run, so a count back at (or near) 0 corroborates "the installer was re-run this morning"; a count still around/above 33 weakens that branch and points instead toward "this session's frontend is linked to a different database than the one checked on 08-27."
3. **Confirm whether a `TBCMSDropboxBridge` process is running anywhere reachable from the affected workstation.** Per the (unmerged) frontend guide, in dev/test this is a manually-started `dotnet run` process (`dropbox-bridge/`, launch profile `http`, port 8088) — it does not run as a service and does not survive a reboot. This session found nothing listening on `localhost:8088` on *this* machine (§5); check the actual target machine the same way (`netstat -ano | findstr 8088`, or browse `http://localhost:8088/api/status`).
4. **If step 1 or 2 shows the placeholder**, update `BridgeUrl` to wherever the real bridge instance is actually running (a `dotnet run` dev instance's `localhost:<port>`, or, once Phase D/E are actually done, the production HTTPS URL) — this is the installer's own documented follow-up step (`Dropbox-Migration-SQL-Install.sql:2405`).
5. **DNS/IIS work for `tbcms-bridge.tatebywater.local` is a Phase D/E production-deployment task**, not something to chase for today's test-environment error — this session already confirmed (§5) that name is NXDOMAIN on the corporate DNS, matching the repo's "Phase D not done" status, and that the resolver itself answered (not a VPN/connectivity problem). Relevant only once an actual bridge server is being stood up per `.docs/bridge-deployment-runbook.md`.
6. **Read `.docs/tbcms-access-frontend-dropbox-update-guide.md` on branch `docs/tbcms-frontend-dropbox-update-guide`** (not on `main` — see §9) before doing any of the above by hand in Access; it has the live, dated `sqlcmd` readings this note draws on in §6, plus the only known documentation of the `frmHome` wiring.
7. **Ask whoever wired `frmHome`'s `Form_Open`** (not in source control — §9) what the actual wrapper code says; it may log more than `StartupBootstrap`'s bare return string, and confirming it would close the one remaining NOT-FOUND in this note.

---

## 9. The dialog's title and closing sentence: not found in this repo

Searched (case-insensitive, tracked and gitignored paths, including `database_assessment/`, `Dropbox-Migration/`, `dropbox-bridge/`, `.docs/`): `"Dropbox session"` as a MsgBox title, `"TBCMS will continue to load"`, and `"document-open operations will not work until this is resolved"`. **None of these three strings exist anywhere in this repository**, tracked or untracked, at the time of this investigation.

The reason is structural, not a search miss: `frmHome`'s `Form_Open`/`Form_Unload` VBA is **not tracked in source control at all**. The only extract of `frmHome` in this repo (`database_assessment/TBCMS/extract/vba/forms/frmHome.txt`) predates the entire Dropbox migration — its `Form_Load` (lines 2044-2054) has no `Form_Open`, no `Form_Unload`, and no mention of Dropbox or `StartupBootstrap` at all, confirming this extract is simply from before this work started, not evidence of what the wiring should be.

The only place in this repo that discusses what should be typed into `frmHome` is, again, the **unmerged** guide (`54e5bc3`, branch `docs/tbcms-frontend-dropbox-update-guide`):

```
"...the currently captured frmHome has only a Form_Load handler and no
Form_Open or Form_Unload at all... There is no tracked source for this
wiring anywhere in the repo — it must be added by hand:"

Private Sub Form_Open(Cancel As Integer)
    DropboxService.StartupBootstrap
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DropboxService.StartupShutdown
End Sub
```

Note this documented snippet calls `StartupBootstrap` as a **bare statement and discards its String return value** — as literally written, it would show no dialog at all on failure. Whoever actually wired the live `frmHome` therefore must have written their own wrapper (something like `result = DropboxService.StartupBootstrap(); If result <> "OK" Then MsgBox result & vbCrLf & vbCrLf & "TBCMS will continue to load..." , vbExclamation, "Dropbox session"`) that is not reflected in this guide either. **Conclusion: the box's title, icon, and closing sentence originate from `frmHome`'s compiled VBA — gitignored per CLAUDE.md's own `.accdb`/`.accde` exclusion — and have no source-control representation to cite beyond this inference.** If a durable fix or message-wording change is ever needed, it has to be made directly in the live Access form, not in either `.bas` file.

---

## 10. A second finding surfaced along the way: the master plan is stale about this entire subsystem

Not asked for directly, but load-bearing for trusting anything else in `.docs/dropbox-migration-plan.md` about Dropbox connectivity: its `▶ NEXT SESSION: START HERE` block (line 16) still reads *"paused 2026-06-04"* and only namechecks the bridge plan as a possible future direction ("Direction may be changing... its Phases A–G are not yet implemented" — line 18). Two distinct gaps follow from that, and they shouldn't be collapsed into one: (1) the bridge was actually designed, implemented, and wired into the default-compiled `DropboxService.bas` in commits dated 2026-06-21 through 2026-06-29 (`01cf2da` through `53ff7d4`) — i.e., within about three weeks of that "paused 2026-06-04" checkpoint, so the plan's status block was already wrong the moment the bridge work landed; and (2) that same status block has now sat unrevised for roughly two more months, through to today (2026-08-28), so it is currently about two months stale relative to the present, on top of having been outdated within weeks of being written. CLAUDE.md inherits the same staleness (its "VBA never builds a path... talks to Dropbox directly via OAuth" description is the pre-bridge/rollback design, not what's compiled by default today). Neither of these was fixed as part of this investigation, per the "do not modify any other files" instruction — flagging here so this staleness isn't silently treated as current by a future reader of either document.

---

## Sources

- `Dropbox-Migration/DropboxService.bas:8-31, 107-115, 461-497, 517-520, 1565-1618, 3926-3947, 4032-4096`
- `Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:2319-2409`
- `.docs/dropbox-bridge-plan.md:1-42, 77-89, 1276-1341`
- `.docs/bridge-deployment-runbook.md:155-230`
- `database_assessment/TBCMS/extract/vba/forms/frmHome.txt:2044-2054`
- `.docs/tbcms-access-frontend-dropbox-update-guide.md` (commit `54e5bc3`, branch `docs/tbcms-frontend-dropbox-update-guide` / `origin/docs/tbcms-frontend-dropbox-update-guide` — **not present on `main`**), especially its Prerequisites §2-§3 (live `sqlcmd` readings) and step 4 (`frmHome` wiring)
- `dropbox-bridge/Properties/launchSettings.json`, `dropbox-bridge/appsettings.json`
- `git log --oneline --all` / `git show` on `01cf2da, da7539d, ddde628, 4852a3d, 57ac812, 0bce1af, e44cb17, e6f66ac, 53ff7d4, 54e5bc3, de31aa7, 2da1417`; `git merge-base --is-ancestor 54e5bc3 main` (result: not an ancestor)
- This session's own probes: `nslookup tbcms-bridge.tatebywater.local` (NXDOMAIN against `TBF-SRVR19.TATEBYWATER.COM`), `curl http://localhost:8088/api/status` (connection failure), `netstat -ano | grep 8088` (no listener) — all run 2026-08-28, read-only, no SQL executed
