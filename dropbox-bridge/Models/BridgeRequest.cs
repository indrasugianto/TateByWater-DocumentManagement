namespace TBCMSDropboxBridge.Models;

// Shared request/response POCOs for the bridge's JSON API. These mirror the
// JSON shapes the VBA DropboxService.bas builds/parses (BridgeRequest helper).
// JSON property names are camelCased by the default System.Text.Json options,
// so VBA reads e.g. "found", "errorSummary", "rawJson", "temporaryLink".

public record MetadataRequest(string Path);
public record MetadataResponse(bool Found, string? ErrorSummary, string? RawJson);

public record FileDownloadLinkRequest(string Path);
public record FileDownloadLinkResponse(string TemporaryLink);

// G34 → bridge-proxy download: VBA POSTs the path, the bridge streams the raw
// bytes back (workstations never reach the Dropbox CDN directly).
public record FileDownloadRequest(string Path);

public record FolderListRequest(string Path);
// response is raw Dropbox JSON string (returned via Results.Text)

public record FolderCreateRequest(string Path);
public record MoveRequest(string FromPath, string ToPath);
public record CopyRequest(string FromPath, string ToPath);
public record DeleteRequest(string Path);

// Upload: path supplied via X-Dropbox-Path header; body = raw file bytes.
public record UploadResponse(bool Success, string? ErrorDetail);

public record StatusResponse(string Status, string? AccountEmail, string? ErrorDetail);
