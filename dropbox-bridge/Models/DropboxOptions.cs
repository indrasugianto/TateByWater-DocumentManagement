namespace TBCMSDropboxBridge.Models;

// Bound from the "Dropbox" section of appsettings.json via
// builder.Services.Configure<DropboxOptions>(...). Injected as
// IOptions<DropboxOptions> into DropboxTokenService and DropboxApiClient.
public sealed class DropboxOptions
{
    public string AppKey { get; set; } = "";
    public string AppSecret { get; set; } = "";
    // Team namespace ID for the shared team tree. Injected as the
    // Dropbox-API-Path-Root header on every call (matches tblDropboxRootConfig).
    public string NamespaceId { get; set; } = "";
    public string RedirectUri { get; set; } = "";
}
