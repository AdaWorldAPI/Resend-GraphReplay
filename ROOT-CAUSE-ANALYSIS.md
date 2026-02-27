# Root Cause Analysis: Empty Folders in IMAP-to-Graph Migration

## Summary

The `Kopano-IMAP-to-Graph-Migration.ps1` script produced **empty folders** in the target Microsoft 365 mailbox. Only the folder structure was created; messages were not migrated.

## Root Causes Identified

### Primary: Broken Custom IMAP Implementation (`SimpleImapClient`)

The script used a hand-rolled `SimpleImapClient` class with multiple critical bugs in IMAP protocol handling:

#### Bug 1: IMAP Literal Parsing in `FetchMessageBytes()` (Line ~560)

```powershell
# BROKEN: Only matches {size} at END of line
if ($line -match '\{(\d+)\}$') {
```

The regex required `{size}` at the **end of the line**, but IMAP servers (including Kopano) may format the FETCH response differently. For example:

```
* 1 FETCH (BODY[] {12345}
```

Some servers include trailing whitespace or additional tokens after the literal size indicator, causing the regex to never match. When the literal size is not detected, `$allBytes` remains empty, and the method returns a zero-length byte array.

#### Bug 2: Mixed StreamReader/Stream Access (Lines ~524-581)

The `FetchMessageBytes()` method accessed the underlying `SslStream` directly for byte-level reading, while other methods used the `StreamReader` wrapper. This caused:

- **Buffering conflicts**: `StreamReader` reads ahead into an internal buffer. When `FetchMessageBytes()` then reads the raw stream, it misses bytes already consumed by the `StreamReader`'s buffer.
- **State corruption**: After `FetchMessageBytes()` reads raw bytes, the `StreamReader` position is out of sync, potentially causing subsequent IMAP commands to fail.

#### Bug 3: No Read Timeout on Raw Stream

The raw byte reading loop had no timeout:

```powershell
while ($bytesRead -lt $literalSize) {
    $chunk = $stream.Read($buffer, $bytesRead, $literalSize - $bytesRead)
    if ($chunk -eq 0) {
        throw "Connection closed while reading literal"
    }
    $bytesRead += $chunk
}
```

On slow connections or when the Kopano server sends data in small chunks, the `Read()` could return 0 without the connection being closed (e.g., temporary network pause), causing the method to throw prematurely.

#### Bug 4: `FetchMessageRaw()` Loses Binary Content (Lines ~499-522)

The fallback `FetchMessageRaw()` method joined response lines with `\r\n`, destroying binary content in MIME messages (attachments, encoded parts). Line-based reading cannot correctly handle binary IMAP literals.

### Secondary: Silent Error Swallowing in `Import-MessageToGraph`

The Graph API import function had three methods, but errors were logged at `Debug` level only:

```powershell
catch {
    Write-Log "MIME import method 1 failed: ..." -Level Debug  # <-- Hidden unless VerboseLogging
}
```

If `FetchMessageBytes()` returned corrupted/empty data, the import would fail silently across all three methods, with no visible error in normal logging. The script would report "0 migrated" without any error indication.

### Tertiary: IMAP Date Format for Kopano

The search criteria used a date format that some Kopano versions don't support:

```powershell
$parts += "SINCE $($StartDate.ToString('dd-MMM-yyyy'))"  # e.g., "01-Jan-2024"
```

Kopano may expect single-digit days (`1-Jan-2024`), causing the SEARCH command to return zero results, making folders appear empty.

## Solution: Replace Custom IMAP with MailKit/MimeKit

Instead of patching the broken custom implementation, the entire `SimpleImapClient` class was replaced with **MailKit** — a battle-tested, production-grade .NET IMAP library.

### What Was Changed

| Component | Before (Broken) | After (Fixed) |
|-----------|-----------------|---------------|
| IMAP client | Custom `SimpleImapClient` class (~300 lines) | MailKit `ImapClient` |
| IMAP literal parsing | Broken regex + raw byte reading | MailKit handles internally |
| SSL/TLS handling | Manual `SslStream` with buffering conflicts | MailKit built-in |
| UTF-7 folder names | Custom `ConvertFrom-ImapUtf7` function | MailKit native support |
| Message fetching | `FetchMessageBytes()` with broken stream access | `folder.GetMessage(uid)` → `MimeMessage` |
| MIME serialization | Raw bytes with corruption risk | `message.WriteTo(stream)` → clean bytes |
| Date extraction | Manual header regex parsing | `message.Date` property |
| Read/Seen flag | Manual `\Seen` string matching | `MessageFlags.Seen` enum |
| Search queries | Manual IMAP command string building | `SearchQuery` fluent API |
| Error handling | Silent `catch` blocks at Debug level | Explicit errors at Warning/Error level |

### New Files

| File | Purpose |
|------|---------|
| `Setup-MailKit.ps1` | Downloads MailKit, MimeKit, BouncyCastle DLLs from NuGet |
| `Kopano-IMAP-Migration-Debug.ps1` | Debug wrapper with all diagnostic options enabled |
| `ROOT-CAUSE-ANALYSIS.md` | This document |

### New Diagnostic Parameters

| Parameter | Purpose |
|-----------|---------|
| `-DiagnosticMode` | Full logging of IMAP capabilities, byte counts, Graph API details |
| `-TestSingleMessage` | Fetch one message, display all details, and stop |
| `-SaveMimeToFile` | Write raw MIME to disk for inspection |
| `-SkipGraphImport` | Test IMAP fetch without needing Graph API credentials |
| `-MimeSavePath` | Custom directory for saved MIME files |

### Deleted Code

- `SimpleImapClient` class (~300 lines) — replaced by MailKit
- `ConvertFrom-ImapUtf7` function — MailKit handles UTF-7 natively
- `Get-ImapClient` function — replaced by `Get-MailKitClient`
- `Get-MessageSearchCriteria` function — replaced by MailKit `SearchQuery` API
- `FetchMessageBytes()` / `FetchMessageRaw()` / `FetchMessageHeaders()` — replaced by MailKit's `GetMessage()`

## Verification Steps

1. Run `Setup-MailKit.ps1` to download dependencies
2. Test IMAP connectivity without Graph API:
   ```powershell
   .\Kopano-IMAP-Migration-Debug.ps1 `
       -ImapServer "imap.elkw.de" `
       -TestSource "user@elkw.de" `
       -TestTarget "user@elkw.de" `
       -TestPassword "xxx" `
       -TenantId "dummy" -ClientId "dummy" -ClientSecret "dummy" `
       -SkipGraphImport
   ```
3. Test single message with full diagnostics:
   ```powershell
   .\Kopano-IMAP-Migration-Debug.ps1 `
       -ImapServer "imap.elkw.de" `
       -TestSource "user@elkw.de" `
       -TestTarget "user@elkw.de" `
       -TestPassword "xxx" `
       -TenantId "xxx" -ClientId "xxx" -ClientSecret "xxx" `
       -TestSingleMessage
   ```
4. Verify output shows `MIME size: <N> bytes` where N > 0
5. Run full migration with a small limit to confirm messages appear in M365
