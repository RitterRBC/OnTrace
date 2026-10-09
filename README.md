# OnTrace

PRESTIGEenterprise OnTrace Logger — a .NET tracing wrapper built on [Apache log4net](https://logging.apache.org/log4net) and [Easy.Logger](https://github.com/NimaAra/EasyLogger).

OnTrace provides a large, easy-to-use static API to write trace/error messages to up to
three independent log targets:

| Target      | Enum value          | Description                                   |
|-------------|---------------------|-----------------------------------------------|
| AppLog      | `enmTarget.AppLog`      | The main application log (file)           |
| EventLog    | `enmTarget.EventLog`    | The Windows/System event log              |
| StartupLog  | `enmTarget.StartupLog`  | A dedicated startup log                   |

Targets can be combined with flags (`AppAndEventLog = 3`, … `AppEventAndStartupLog = 7`).

## Target frameworks

`net452`, `net462`, `net472`.

## Public API (summary)

- `TraceError` / `TraceErrorFormat`
- `TraceWarning` / `TraceWarningFormat`
- `TraceInformation` / `TraceInformationFormat`
- `TraceFatal` / `TraceFatalFormat`
- `WriteLine` / `WriteLineFormat` / `WriteLineLevel` / `WriteLineLevelFormat`
- `TraceException` / `TraceFatalException` / `SetError` / `GetErrorStr`
- Dynamic file logging: `TraceError(…, path, file)` style overloads plus `TraceClear(path, file)`
- `Reconfigure(...)`, `PatchAppenders()`, `GetLogFileName(...)`
- Properties: `AppLoggerEnabled`, `EventLoggerEnabled`, `StartupLoggerEnabled`, `TraceLevel`, `IsLicenseCheckEnabled`

Every group provides several overloads for a raw message, a formatted message (`{0}`, `{1}`, …)
and a `params object[]` form, plus variants that select a `TraceLevel` and a target.

## Trace levels

`enmTraceLevel` is a bit mask:

| Level       | Value | Level       | Value |
|-------------|-------|-------------|-------|
| `Level_1`   | 1     | `Level_6`   | 32    |
| `Level_2`   | 2     | `Level_7`   | 64    |
| `Level_3`   | 4     | `Level_8`   | 128   |
| `Level_4`   | 8     | `Level_9`   | 256   |
| `Level_5`   | 16    | `Level_All` | 511   |

Set `OnTrace.TraceLevel` (or the `DefaultTraceLevel` setting in `app.config`/`web.config`) to a
combination of these bits. `TraceInformation`, `TraceWarning` and `WriteLine*` messages are only
written when the requested bit is enabled. **`TraceError` and `TraceFatal` messages are always
written, independent of the configured trace level.**

## Configuration

OnTrace is configured from the host application's `.config` file (the `<log4net>` section plus a
few application settings). The three loggers are created from these keys (all optional):

| AppSetting                | Default       | Purpose                                    |
|---------------------------|---------------|--------------------------------------------|
| `AppLoggerSelectLogger`   | `"AppLogger"`   | Logger name for the application log    |
| `EventLoggerSelectLogger` | `"EventLogger"` | Logger name for the event log         |
| `StartupLoggerSelectLogger` | `"StartupLogger"` | Logger name for the startup log    |
| `DefaultTraceLevel`       | (none)        | Initial value of `OnTrace.TraceLevel`      |
| `AppLogger` / `EventLogger` / `StartupLogger` | (none) | `true`/`false` to enable each logger |
| `Client`                  | `"Standard"`  | Placeholder used when patching file paths  |
| `BasePathLogFile`         | `%PROGRAMDATA%\Online Software AG\Logs\` | Base folder for log files |

OnTrace expects the loggers to have appenders. `GetLogFileName()` reads the following appender
names by default: `LogFileAppender` (AppLog), `EventlogAppender` (EventLog) and
`StartupLogFileAppender` (StartupLog).

See [sample-log4net.config](OnTrace/sample-log4net.config) for a complete example.

## Dynamic file logging

The `(…, path, fileName)` overloads write directly to an external file (creating a dynamic
`RollingFileAppender` with an `AsyncBufferingForwardingAppender`). Call `TraceClear(path, fileName)`
when the file usage is finished so the appenders are closed and released.

## Examples

```vb
OnTrace.TraceInformation("Application started")
OnTrace.TraceErrorFormat("User {0} failed to log in after {1} tries", user, tries)
OnTrace.TraceWarning("Disk nearly full", OnTrace.enmTraceLevel.Level_3)
OnTrace.TraceFatal(ex, "Fatal condition", tempDir, "crash.log")
OnTrace.TraceError("write to a specific file", OnTrace.enmTraceLevel.Level_All, tempDir, "detail.log")
OnTrace.TraceClear(tempDir, "detail.log")
```

## Building and testing

```
dotnet restore OnTrace.sln
dotnet build  OnTrace.sln -c Release
dotnet test   OnTrace.Tests/OnTrace.Tests.csproj -c Release
dotnet pack   OnTrace/OnTrace.vbproj -c Release -o ./artifacts
```

The unit tests cover level filtering, the "errors always written" behaviour and target routing
using in-memory appenders.

## Packaging

The NuGet package is produced with `dotnet pack` (no hand-maintained `.nuspec`). Versioning,
dependencies, license and description are driven by `OnTrace/OnTrace.vbproj`.

## CI

`.github/workflows/ci.yml` restores, builds, tests and packs on every push/PR (Windows).

## Licensing

OnTrace is a wrapper around Apache log4net. The log4net license and notice are shipped alongside
the assembly (`LICENSE_L4N.txt`, `NOTICE_L4N.txt`) and inside the NuGet package.
