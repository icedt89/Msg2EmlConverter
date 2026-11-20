# 📬 Msg2EmlConverter

Convert Outlook **MSG ↔ EML** files with a simple Windows command-line tool.
Drag a message onto the EXE or pass its path from the command line.

> 🤖 This application was entirely vibe-coded with AI.

## 🚀 Quick start

If you already have the standalone `Msg2EmlConverter.exe`:

1. Place it in a folder you can write to.
2. Drag **one** `.msg` or `.eml` file from File Explorer onto the EXE.
3. Find the converted file **next to the EXE**, with the same base name and the
   other extension: `message.msg` → `message.eml`, or the reverse.

The application exits automatically after a successful conversion.
To create the standalone EXE from source, follow [Build](#-build).

You can also run it from PowerShell in the folder containing the EXE:

```powershell
.\Msg2EmlConverter.exe "C:\Mail\My message.msg"
.\Msg2EmlConverter.exe "C:\Mail\My message.eml"
```

> ⚠️ **Existing output files are overwritten.** Output always goes next to the
> application, regardless of the input file's folder or the working directory.
> Messages with the same base name therefore share the same output path.

## ✨ Features

- Convert messages in both directions: MSG to EML and EML to MSG.
- Preserve plain text and HTML bodies, Unicode text, sender and recipient
  information, transport headers, and message timestamps.
- Handle file attachments, embedded messages, and inline resources using
  Content-ID or relative and absolute Content-Location.
- Preserve attachment MIME types, declared text character sets, and existing
  creation and modification timestamps.
- Keep attached multipart content together, separate from the message body.
- Run as a standalone Windows x64 EXE without a separate .NET installation.

## 📋 Requirements

| Task | Requirements |
| --- | --- |
| Run the standalone EXE | Windows x64 and a writable application folder; no separate .NET installation |
| Run a regular build | .NET 10 Runtime and a writable application folder |
| Build or run the integration tests | .NET 10 SDK; current publishing and test commands target Windows |

`Source/global.json` selects SDK `10.0.100` or a later stable .NET 10 SDK allowed by its
`latestMinor` roll-forward policy. Cross-platform operation is not verified.

## 🔨 Build

Run the build and development commands below from the `Source` directory.
From the repository root, enter it first:

```powershell
Set-Location .\Source
```

### Standalone Windows x64 EXE

Double-click `publish.cmd` or run it from PowerShell:

```powershell
.\publish.cmd
```

The script creates a Release build with the .NET Runtime and dependencies
bundled into `publish\Msg2EmlConverter.exe`. You can distribute this EXE on its
own; any generated PDB file is optional. The script reports success or failure
and waits for a key press.

### Regular build

```powershell
dotnet build Msg2EmlConverter/Msg2EmlConverter.csproj -c Release
& ".\Msg2EmlConverter\bin\Release\net10.0\Msg2EmlConverter.exe" "C:\Mail\My message.msg"
```

When copying a regular build, include the other files in its output directory.
The target computer needs the .NET 10 Runtime.

## 🧭 Behavior and limitations

- Each invocation accepts exactly one existing file with a `.msg` or `.eml`
  extension, case-insensitively. Batch conversion and an output-folder option
  are not available.
- Embedded MSG messages become EML attachments. During EML-to-MSG conversion,
  embedded EML messages remain EML file attachments with their own headers and
  attachments.
- Duplicate attachment names receive numbered names when necessary. Missing
  attachment timestamps are not added.
- The supported fields and attachment scenarios are covered by generated test
  messages. Complete preservation of every Outlook-specific property or a
  byte-for-byte round trip is not guaranteed. Signed or encrypted messages,
  calendar items, contacts, and tasks are outside the current test coverage.

## 🔍 Logs and troubleshooting

The application writes colored status and debug messages to stdout. Errors and
full exception stack traces appear in red on stderr.

| Exit code | Meaning |
| --- | --- |
| `0` | Conversion completed successfully |
| `1` | Invalid arguments, a missing file, or a conversion error |

After an error, press Enter to close the console window. If standard input is
redirected, the application exits without waiting.

If conversion fails, check that the input file exists, its extension is `.msg`
or `.eml`, and the application folder is writable.

Logs can contain message subjects, email addresses, headers, file paths, and
attachment names. Review and redact them before sharing. For a bug report,
include the application version or source revision, conversion direction,
error output, and, where possible, an anonymized sample that reproduces the
problem.

## 🛠️ Development

### Run integration tests

Run this command from the `Source` directory:

```powershell
dotnet run --project Msg2EmlConverter.Tests
```

The tests generate MSG and EML files and exercise conversion in both
directions. They cover output location and overwriting, invalid input, plain
text and HTML with umlauts, transport headers, recipients and recipient
groups, sending on behalf of another sender, timestamps, file attachments,
inline images, embedded messages, and duplicate attachment names. Regression
checks also cover multipart attachments, Content-Location, MIME types,
attachment timestamps, and text attachment character sets.

The test project is a console application; use the command above to execute
its checks.

### Use the converters directly

The static converters can also be called from C#. Register code pages once
at startup and dispose of returned messages with `using`:

```csharp
using System.Text;
using MimeKit;
using MsgReader.Outlook;
using Msg2EmlConverter;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
using var originalMsg = new Storage.Message("message.msg");
using var eml = MsgToEmlConverter.Convert(originalMsg);
using var msg = EmlToMsgConverter.Convert(eml);
msg.Save("converted.msg");
```

The EML-to-MSG converter uses MsgKit and returns an independent
`Storage.Message`. The command-line application registers code pages for you.

### Logging configuration

Serilog configuration, including the `Debug` minimum level and console themes,
is defined in `Source/Msg2EmlConverter/Program.cs`. No `appsettings.json` is required.
