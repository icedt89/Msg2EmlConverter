using System.Text;
using Msg2EmlConverter;
using MimeKit;
using MsgReader.Outlook;
using Serilog;
using Serilog.Events;
using Serilog.Sinks.SystemConsole.Themes;

// Preserve Unicode log messages when output is redirected.
Console.OutputEncoding = Encoding.UTF8;

// Older mail can use Windows-1252 and other encodings that modern .NET does not enable by default.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

// Keep the entire error message and stack trace red.
// ConsoleThemeStyle contains aliases, so duplicate numeric values must be removed before building the dictionary.
var errorTheme = new AnsiConsoleTheme(Enum.GetValues<ConsoleThemeStyle>()
    .Distinct()
    .ToDictionary(style => style, _ => "\e[31m"));

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    // Exclude errors from the normal sink so they appear once, entirely red, on stderr.
    .WriteTo.Logger(log => log
        .Filter.ByExcluding(logEvent => logEvent.Level >= LogEventLevel.Error)
        .WriteTo.Console(theme: AnsiConsoleTheme.Code))
    .WriteTo.Console(
        restrictedToMinimumLevel: LogEventLevel.Error,
        standardErrorFromLevel: LogEventLevel.Error,
        theme: errorTheme)
    .CreateLogger();

Log.Debug("Application directory: {ApplicationDirectory:l}. Received {ArgumentCount} command-line argument(s).",
    AppContext.BaseDirectory, args.Length);

if (args.Length != 1)
{
    ReportErrorAndWait("Usage: Msg2EmlConverter.exe \"path to message.msg or message.eml\"");

    return 1;
}

try
{
    if (!TryGetInputMessagePath(args[0], out var inputPath))
    {
        return 1;
    }

    var outputPath = string.Equals(Path.GetExtension(inputPath), ".msg", StringComparison.OrdinalIgnoreCase)
        ? ConvertMsgFileToEml(inputPath)
        : ConvertEmlFileToMsg(inputPath);

    Log.Information("{OutputType} file saved to {OutputPath:l}.", Path.GetExtension(outputPath).TrimStart('.').ToUpperInvariant(), outputPath);

    return 0;
}
catch (Exception ex)
{
    ReportErrorAndWait("Conversion failed.", ex);

    return 1;
}

static void ReportErrorAndWait(string message, Exception? exception = null)
{
    // Passing the exception separately preserves its type, inner exceptions and full stack trace.
    Log.Error(exception, "{ErrorMessage}", message);

    // Keep interactive error messages visible; redirected input must not block automation.
    if (!Console.IsInputRedirected)
    {
        Log.Information("Press Enter to exit.");

        Console.ReadLine();
    }
    else
    {
        Log.Debug("Input is redirected; exiting without waiting for Enter.");
    }
}

static bool TryGetInputMessagePath(string argument, out string inputPath)
{
    inputPath = Path.GetFullPath(argument);
    Log.Debug("Resolved input path: {InputPath:l}.", inputPath);
    var extension = Path.GetExtension(inputPath);
    if (!string.Equals(extension, ".msg", StringComparison.OrdinalIgnoreCase)
        && !string.Equals(extension, ".eml", StringComparison.OrdinalIgnoreCase))
    {
        ReportErrorAndWait("The input file must have the .msg or .eml extension.");

        return false;
    }

    if (!File.Exists(inputPath))
    {
        ReportErrorAndWait($"File not found: {inputPath}");

        return false;
    }

    return true;
}

static string ConvertMsgFileToEml(string inputMsgPath)
{
    var outputEmlPath = GetOutputPath(inputMsgPath, ".eml");
    Log.Information("Reading MSG input {InputPath:l}.", inputMsgPath);
    using var msg = new Storage.Message(inputMsgPath);
    using var eml = MsgToEmlConverter.Convert(msg);

    Log.Information("Writing EML output {OutputPath:l}.", outputEmlPath);
    eml.WriteTo(outputEmlPath);

    return outputEmlPath;
}

static string ConvertEmlFileToMsg(string inputEmlPath)
{
    var outputMsgPath = GetOutputPath(inputEmlPath, ".msg");
    Log.Information("Reading EML input {InputPath:l}.", inputEmlPath);
    using var eml = MimeMessage.Load(inputEmlPath);
    using var msg = EmlToMsgConverter.Convert(eml);

    Log.Information("Writing MSG output {OutputPath:l}.", outputMsgPath);
    msg.Save(outputMsgPath);

    return outputMsgPath;
}

static string GetOutputPath(string inputPath, string outputExtension)
{
    // Explorer can supply files from any directory; outputs always belong beside the executable.
    return Path.Combine(AppContext.BaseDirectory, Path.ChangeExtension(Path.GetFileName(inputPath), outputExtension));
}
