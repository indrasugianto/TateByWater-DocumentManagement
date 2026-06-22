namespace TBCMSDropboxBridge.Models;

// Classifies a Dropbox API error so the global exception handler (Program.cs)
// can map it to the HTTP status the VBA caller expects (see plan B.7):
//   NotFound  -> 404   (Dropbox returns HTTP 409 path/not_found)
//   Conflict  -> 409   (Dropbox to/conflict etc. — lets MoveFile/CopyFile tell
//                       "already exists at destination" apart from a failure)
//   AuthError -> 401   (revoked/invalid token)
//   RateLimited -> 429 (honor Retry-After; see G33)
//   Other     -> 500
public enum DropboxErrorKind
{
    NotFound,
    Conflict,
    AuthError,
    RateLimited,
    Other
}

// Thrown by DropboxApiClient when a Dropbox call fails. Carries the original
// error_summary so it can be logged/surfaced, and a Kind the error handler maps
// to an HTTP status. Translating Dropbox's uniform "HTTP 409 + error_summary"
// error model into distinct statuses here is what makes the B.7 contract real
// (otherwise everything non-503/403 collapses to 500).
public sealed class DropboxApiException : Exception
{
    public DropboxErrorKind Kind { get; }
    public int DropboxStatus { get; }
    public string? ErrorSummary { get; }
    public int? RetryAfterSeconds { get; }

    public DropboxApiException(
        DropboxErrorKind kind,
        int dropboxStatus,
        string? errorSummary,
        string message,
        int? retryAfterSeconds = null)
        : base(message)
    {
        Kind = kind;
        DropboxStatus = dropboxStatus;
        ErrorSummary = errorSummary;
        RetryAfterSeconds = retryAfterSeconds;
    }
}
