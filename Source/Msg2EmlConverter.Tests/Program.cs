using System.Diagnostics;
using System.Text;
using Msg2EmlConverter;
using MimeKit;
using OpenMcdf;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

// Integration tests without an additional test framework: create actual MSG files,
// run the application as a separate process, and read the resulting EML files.
string testDirectory = Path.Combine(AppContext.BaseDirectory, "test-data", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(testDirectory);
int checks = 0;

Check(Run().Code != 0, "Fehlendes Argument wird abgewiesen");
Check(Run("eins.msg", "zwei.msg").Code != 0, "Zusätzliche Argumente werden abgewiesen");
Check(Run("fehlt.msg").Code != 0, "Fehlende Datei wird abgewiesen");
Check(Run("falsch.txt").Code != 0, "Falsche Endung wird abgewiesen");
File.WriteAllText(Path.Combine(testDirectory, "kaputt.msg"), "keine MSG-Datei");
var invalid = Run("kaputt.msg");
Check(invalid.Code != 0 && invalid.Error.Contains("Conversion failed"), "Beschädigte MSG liefert eine verständliche Fehlermeldung");
Check(invalid.Error.Contains("Exception:") && invalid.Error.Contains("ConvertMsgFileToEml"), "Fehler enthält Exception-Typ und Stacktrace");
Check(invalid.Error.Split(" ERR]").Length == 2 && !invalid.Output.Contains("Conversion failed"), "Serilog schreibt Fehler genau einmal auf stderr");
Check(!File.Exists(Path.Combine(AppContext.BaseDirectory, "kaputt.eml")), "Ungültige MSG erzeugt keine Ausgabe");
Check(Run("fehlt.eml").Code != 0, "Fehlende EML wird abgewiesen");
File.WriteAllBytes(Path.Combine(testDirectory, "leer.eml"), Array.Empty<byte>());
var emptyEml = Run("leer.eml");
Check(emptyEml.Code != 0 && emptyEml.Error.Contains("Conversion failed"), "Leere EML liefert einen Konvertierungsfehler");
Check(!File.Exists(Path.Combine(AppContext.BaseDirectory, "leer.msg")), "Ungültige EML erzeugt keine MSG");

using (var minimal = RootStorage.Create(Path.Combine(testDirectory, "minimal.msg")))
{
    WriteProperties(minimal, 32,
        (0x001A001F, Encoding.Unicode.GetBytes("IPM.Note\0")),
        (0x0037001F, Encoding.Unicode.GetBytes("Minimal\0")),
        (0x1000001F, Encoding.Unicode.GetBytes("Nur Text\0")));
}
Check(Run("minimal.msg").Code == 0, "MSG ohne optionale Metadaten wird konvertiert");
using (var minimalEml = MimeMessage.Load(Path.Combine(AppContext.BaseDirectory, "minimal.eml")))
{
    Check(minimalEml.Subject == "Minimal" && minimalEml.TextBody?.TrimEnd('\r', '\n') == "Nur Text", "Minimale MSG behält Betreff und Text");
}

foreach (var (withHeaders, delegated) in new[] { (false, false), (true, false), (false, true), (true, true) })
{
    string fileName = $"Nachricht mit Umlauten ä {withHeaders} {delegated}.MSG";
    using (var root = RootStorage.Create(Path.Combine(testDirectory, fileName)))
    {
        WriteMessage(root, withHeaders, delegated: delegated);
    }

    string output = Path.Combine(AppContext.BaseDirectory, Path.ChangeExtension(fileName, ".eml"));
    using (var msg = new MsgReader.Outlook.Storage.Message(Path.Combine(testDirectory, fileName)))
    using (var converted = MsgToEmlConverter.Convert(msg))
    {
        Check(converted.Subject == "Prüfung äöü" && converted.TextBody?.TrimEnd('\0', '\r', '\n') == "Hallo Grüße!", "Konverter kann direkt verwendet werden");
        Check(converted.Attachments.OfType<MessagePart>().Single().Message?.Subject == "Verschachtelt", "Direkte Konvertierung übernimmt eingebettete Nachrichten");
    }
    File.WriteAllText(output, "Alte EML");
    var result = Run(withHeaders ? Path.Combine(testDirectory, fileName) : fileName);
    Check(result.Code == 0, $"Konvertierung erfolgreich: {result.Error}");
    Check(result.Output.Contains("EML file saved to") && result.Output.Contains(output) && result.Error.Length == 0, $"Serilog protokolliert den Ausgabepfad auf stdout. Erwartet: {output}\nstdout: {result.Output}\nstderr: {result.Error}");
    Check(result.Output.Contains("Converting MSG message") && result.Output.Contains("converted successfully"), "Konsolenlog enthält Konvertierungsstart und erfolgreichen Abschluss");
    Check(result.Output.Contains(" DBG]") && result.Output.Contains("Adding file attachment") && result.Output.Contains("Adding inline resource") && result.Output.Contains("Converting embedded MSG attachment"), "Debug-Logs beschreiben die Verarbeitung der Anhänge");
    Check(File.Exists(output) && !File.Exists(Path.ChangeExtension(Path.Combine(testDirectory, fileName), ".eml")), "EML liegt neben der Anwendung, mit gleichem Basisnamen");
    using var eml = MimeMessage.Load(output);
    Check(eml.Subject == "Prüfung äöü", "Betreff bleibt erhalten");
    Check(eml.TextBody?.TrimEnd('\0', '\r', '\n') == "Hallo Grüße!", "Text und Umlaute bleiben erhalten");
    Check(eml.HtmlBody?.TrimEnd('\0', '\r', '\n') == "<html><body>Grüße<img src=\"cid:bild@example.org\"></body></html>", "HTML bleibt erhalten");
    Check(eml.From.Mailboxes.Single().Address == (delegated ? "represented@example.org" : "from@example.org"), "Absender bleibt erhalten");
    Check(!delegated || eml.Sender?.Address == "from@example.org", "Versand im Auftrag bleibt korrekt erhalten");
    Check(eml.To.Mailboxes.Single().Address == "to@example.org", "To bleibt erhalten");
    Check(eml.Cc.Mailboxes.Single().Address == "cc@example.org", "Cc bleibt erhalten");
    Check(eml.Bcc.Mailboxes.Single().Address == "bcc@example.org", "Bcc wird auch bei vorhandenen To-Headern ergänzt");
    Check(eml.Date == new DateTimeOffset(2024, 2, 3, 10, 20, 30, TimeSpan.Zero), "Sendedatum bleibt erhalten");
    Check(eml.MessageId == "message@example.org" && eml.InReplyTo == "previous@example.org", "Message-ID und In-Reply-To bleiben erhalten");
    Check(eml.References.SequenceEqual(new[] { "first@example.org", "previous@example.org" }), "References bleiben erhalten");
    Check(eml.Importance == MessageImportance.High && eml.Headers[HeaderId.Sensitivity] == "Private", "Wichtigkeit und Vertraulichkeit bleiben erhalten");
    Check(eml.Headers.Contains(HeaderId.DispositionNotificationTo) && eml.Headers.Contains(HeaderId.ReturnReceiptTo), "Bestätigungsanforderungen bleiben erhalten");
    var file = eml.Attachments.OfType<MimePart>().Single();
    using var data = new MemoryStream();
    file.Content!.DecodeTo(data);
    Check(file.FileName == "Anhang.txt" && Encoding.UTF8.GetString(data.ToArray()) == "Anhang Grüße", "Dateianhang bleibt unverändert");
    Check(eml.BodyParts.OfType<MimePart>().Any(part => part.ContentId == "bild@example.org" && !part.IsAttachment), "Inline-Bild bleibt verknüpft");
    var nested = eml.Attachments.OfType<MessagePart>().Single();
    Check(nested.Message?.Subject == "Verschachtelt" && nested.Message.TextBody?.TrimEnd('\0', '\r', '\n') == "Innere Nachricht", "Eingebettete MSG wird als EML-Anhang übernommen");
    if (withHeaders)
    {
        Check(eml.Headers["X-Test"] == "eins" && eml.ReplyTo.Mailboxes.Single().Address == "reply@example.org", "Eigene Header und Reply-To bleiben erhalten");
        Check(eml.Headers.Count(header => header.Field == "Received") == 2, "Wiederholte Header bleiben erhalten");
    }

    using (var directMsg = EmlToMsgConverter.Convert(eml))
    using (var directEml = MsgToEmlConverter.Convert(directMsg))
    {
        Check(directEml.Subject == eml.Subject && directEml.TextBody?.TrimEnd('\r', '\n') == "Hallo Grüße!", "EML-Konverter kann direkt verwendet werden");
    }

    var reverseName = $"Rückweg ä {withHeaders} {delegated}.EML";
    var reverseInput = Path.Combine(testDirectory, reverseName);
    eml.WriteTo(reverseInput);
    var reverseOutput = Path.Combine(AppContext.BaseDirectory, Path.ChangeExtension(reverseName, ".msg"));
    File.WriteAllText(reverseOutput, "Alte MSG");
    var reverseResult = Run(withHeaders ? reverseInput : reverseName);
    Check(reverseResult.Code == 0, $"EML nach MSG erfolgreich: {reverseResult.Error}");
    Check(reverseResult.Output.Contains("Converting EML message") && reverseResult.Output.Contains("MSG file saved to"), "EML-Konvertierung wird auf Englisch protokolliert");
    Check(File.Exists(reverseOutput) && !File.Exists(Path.ChangeExtension(reverseInput, ".msg")), "MSG liegt mit gleichem Basisnamen neben der EXE");

    using var restoredMsg = new MsgReader.Outlook.Storage.Message(reverseOutput);
    using var restored = MsgToEmlConverter.Convert(restoredMsg);
    Check(restored.Subject == eml.Subject, "EML nach MSG erhält den Betreff");
    Check(restored.TextBody?.TrimEnd('\r', '\n') == "Hallo Grüße!", "EML nach MSG erhält Text und Umlaute");
    Check(restored.HtmlBody?.TrimEnd('\r', '\n') == eml.HtmlBody?.TrimEnd('\r', '\n'), "EML nach MSG erhält HTML");
    Check(restored.From.Mailboxes.Single().Address == eml.From.Mailboxes.Single().Address, "EML nach MSG erhält From");
    Check(!delegated || restored.Sender?.Address == "from@example.org", "EML nach MSG erhält Versand im Auftrag");
    Check(restored.To.Mailboxes.Single().Address == "to@example.org" && restored.Cc.Mailboxes.Single().Address == "cc@example.org" && restored.Bcc.Mailboxes.Single().Address == "bcc@example.org", "EML nach MSG erhält To, Cc und Bcc");
    Check(restored.Date == eml.Date && restored.MessageId == eml.MessageId, "EML nach MSG erhält Datum und Message-ID");
    Check(restored.InReplyTo == eml.InReplyTo && restored.References.SequenceEqual(eml.References), "EML nach MSG erhält Threading");
    Check(restored.Importance == eml.Importance && restored.Headers[HeaderId.Sensitivity] == "Private", "EML nach MSG erhält Wichtigkeit und Vertraulichkeit");
    Check(restored.Headers.Contains(HeaderId.DispositionNotificationTo) && restored.Headers.Contains(HeaderId.ReturnReceiptTo), "EML nach MSG erhält Bestätigungsanforderungen");
    var restoredFile = restored.Attachments.OfType<MimePart>().Single(part => part.FileName == "Anhang.txt");
    using var restoredData = new MemoryStream();
    restoredFile.Content!.DecodeTo(restoredData);
    Check(Encoding.UTF8.GetString(restoredData.ToArray()) == "Anhang Grüße", "EML nach MSG erhält Dateianhänge unverändert");
    Check(restored.BodyParts.OfType<MimePart>().Any(part => part.ContentId == "bild@example.org" && !part.IsAttachment), "EML nach MSG erhält Inline-Bilder");
    var attachedMessage = restored.Attachments.OfType<MessagePart>().Single().Message;
    Check(attachedMessage?.Subject == "Verschachtelt" && attachedMessage.TextBody?.TrimEnd('\r', '\n') == "Innere Nachricht", "Eingebettete EML bleibt als vollständiger Dateianhang erhalten");
    if (withHeaders)
    {
        Check(restored.ReplyTo.Mailboxes.Single().Address == "reply@example.org" && restored.Headers["X-Test"] == "eins", "EML nach MSG erhält Reply-To und eigene Header");
        Check(restored.Headers.Count(header => header.Field == "Received") == 2, "EML nach MSG erhält wiederholte Header");
    }
}

using (var nativeEml = new MimeMessage())
{
    nativeEml.Subject = "Unabhängige EML ä";
    nativeEml.From.Add(new MailboxAddress("Absender", "author@example.org"));
    nativeEml.Sender = new MailboxAddress("Vertreter", "delegate@example.org");
    var group = new GroupAddress("Team");
    group.Members.Add(new MailboxAddress("Person 1", "one@example.org"));
    group.Members.Add(new MailboxAddress("Person 2", "two@example.org"));
    nativeEml.To.Add(group);
    nativeEml.Importance = MessageImportance.Low;
    nativeEml.Priority = MessagePriority.NonUrgent;
    nativeEml.Date = new DateTimeOffset(2023, 4, 5, 12, 13, 14, TimeSpan.FromHours(2));
    var builder = new BodyBuilder { TextBody = "Eigener Nachrichtentext äöü" };
    builder.Attachments.Add("notes.txt", Encoding.UTF8.GetBytes("Erster Anhang"), ContentType.Parse("text/plain"));
    builder.Attachments.Add("notes.txt", Encoding.UTF8.GetBytes("Zweiter Anhang"), ContentType.Parse("text/plain"));
    var inline = builder.LinkedResources.Add("logo.png", new byte[] { 7, 8, 9 });
    inline.ContentId = "logo@example.org";
    builder.HtmlBody = "<html><body>Logo<img src=\"cid:logo@example.org\"></body></html>";
    nativeEml.Body = builder.ToMessageBody();
    using var nativeMsg = EmlToMsgConverter.Convert(nativeEml);
    using var converted = MsgToEmlConverter.Convert(nativeMsg);
    Check(converted.TextBody?.TrimEnd('\r', '\n') == builder.TextBody, "Textanhang ersetzt nicht den Nachrichtentext");
    Check(converted.From.Mailboxes.Single().Address == "author@example.org" && converted.Sender?.Address == "delegate@example.org", "Eigenständige EML erhält From und Sender");
    Check(nativeMsg.Recipients.Count == 2, "Empfängergruppen werden in MSG-Empfänger aufgelöst");
    Check(converted.To.Mailboxes.Select(mailbox => mailbox.Address).SequenceEqual(new[] { "one@example.org", "two@example.org" }), "Gruppen bleiben beim Rückweg im Header erhalten");
    Check(converted.Importance == MessageImportance.Low && converted.Priority == MessagePriority.NonUrgent, "Niedrige Wichtigkeit und Priorität bleiben erhalten");
    Check(converted.Date == nativeEml.Date, "Zeitzonen werden korrekt übernommen");
    var attachments = converted.Attachments.OfType<MimePart>().ToList();
    Check(attachments.Count == 2 && attachments.Select(part => part.FileName).Distinct().Count() == 2, "Gleichnamige Anhänge werden unter eindeutigen Namen erhalten");
    var attachmentTexts = new List<string>();
    foreach (var attachment in attachments)
    {
        using var bytes = new MemoryStream();
        attachment.Content!.DecodeTo(bytes);
        attachmentTexts.Add(Encoding.UTF8.GetString(bytes.ToArray()));
    }
    Check(attachmentTexts.SequenceEqual(new[] { "Erster Anhang", "Zweiter Anhang" }), "Gleichnamige Anhänge behalten ihre jeweiligen Daten");
    var restoredInline = converted.BodyParts.OfType<MimePart>().Single(part => part.ContentId == "logo@example.org");
    using var inlineData = new MemoryStream();
    restoredInline.Content!.DecodeTo(inlineData);
    Check(inlineData.ToArray().SequenceEqual(new byte[] { 7, 8, 9 }), "Inline-Bilddaten bleiben unverändert");
}

using (var emptyBodyEml = new MimeMessage())
{
    emptyBodyEml.Subject = "Ohne Absender und Inhalt";
    using var emptyBodyMsg = EmlToMsgConverter.Convert(emptyBodyEml);
    Check(emptyBodyMsg.Subject == emptyBodyEml.Subject, "EML ohne Absender und Nachrichtentext wird konvertiert");
}

ConverterRegressionTests.Run(testDirectory, Check);

Console.WriteLine($"Alle {checks} Prüfungen erfolgreich.");

(int Code, string Error, string Output) Run(params string[] arguments)
{
    var startInfo = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "Msg2EmlConverter.exe"))
    {
        WorkingDirectory = testDirectory,
        RedirectStandardInput = true,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        StandardOutputEncoding = Encoding.UTF8,
        StandardErrorEncoding = Encoding.UTF8,
        UseShellExecute = false,
        CreateNoWindow = true
    };
    foreach (string argument in arguments)
    {
        startInfo.ArgumentList.Add(argument);
    }
    using var process = Process.Start(startInfo)!;
    process.StandardInput.Close();
    var stdout = process.StandardOutput.ReadToEndAsync();
    string error = process.StandardError.ReadToEnd();
    process.WaitForExit();
    string output = stdout.GetAwaiter().GetResult();
    return (process.ExitCode, error, output);
}

void Check(bool condition, string message)
{
    if (!condition)
    {
        throw new InvalidOperationException(message);
    }
    checks++;
}

static void WriteMessage(Storage storage, bool withHeaders, bool embedded = false, bool delegated = false)
{
    var properties = new List<(uint Tag, byte[] Value)>
    {
        (0x001A001F, Encoding.Unicode.GetBytes("IPM.Note\0")),
        (0x0037001F, Encoding.Unicode.GetBytes(embedded ? "Verschachtelt\0" : "Prüfung äöü\0")),
        (0x1000001F, Encoding.Unicode.GetBytes(embedded ? "Innere Nachricht\0" : "Hallo Grüße!\0")),
        (0x0C1A001F, Encoding.Unicode.GetBytes("Absender\0")),
        (0x0C1E001F, Encoding.Unicode.GetBytes("SMTP\0")),
        (0x0C1F001F, Encoding.Unicode.GetBytes("from@example.org\0")),
        (0x00390040, BitConverter.GetBytes(new DateTime(2024, 2, 3, 10, 20, 30, DateTimeKind.Utc).ToFileTimeUtc())),
        (0x1035001F, Encoding.Unicode.GetBytes("<message@example.org>\0")),
        (0x1042001F, Encoding.Unicode.GetBytes("<previous@example.org>\0")),
        (0x1039001F, Encoding.Unicode.GetBytes("<first@example.org> <previous@example.org>\0")),
        (0x00170003, BitConverter.GetBytes(2)),
        (0x00360003, BitConverter.GetBytes(2)),
        (0x0029000B, BitConverter.GetBytes((short)1)),
        (0x0023000B, BitConverter.GetBytes((short)1))
    };
    if (!embedded)
    {
        properties.Add((0x10130102, Encoding.UTF8.GetBytes("<html><body>Grüße<img src=\"cid:bild@example.org\"></body></html>")));
    }
    if (delegated)
    {
        properties.Add((0x0042001F, Encoding.Unicode.GetBytes("Vertreten\0")));
        properties.Add((0x0064001F, Encoding.Unicode.GetBytes("SMTP\0")));
        properties.Add((0x0065001F, Encoding.Unicode.GetBytes("represented@example.org\0")));
    }
    if (withHeaders)
    {
        properties.Add((0x007D001F, Encoding.Unicode.GetBytes($"From: Absender <{(delegated ? "represented" : "from")}@example.org>\r\nTo: to@example.org\r\nSubject: Prüfung äöü\r\nDate: Sat, 3 Feb 2024 10:20:30 +0000\r\nReply-To: reply@example.org\r\nReceived: eins\r\nReceived: zwei\r\nX-Test: eins\r\nContent-Type: multipart/mixed; boundary=alt\r\n\0")));
    }
    WriteProperties(storage, embedded ? 24 : 32, properties.ToArray());
    if (embedded)
    {
        return;
    }

    for (int index = 0; index < 3; index++)
    {
        var recipient = storage.CreateStorage($"__recip_version1.0_#{index:X8}");
        string address = new[] { "to", "cc", "bcc" }[index] + "@example.org\0";
        WriteProperties(recipient, 8,
            (0x0C150003, BitConverter.GetBytes(index + 1)),
            (0x3001001F, Encoding.Unicode.GetBytes(address)),
            (0x3002001F, Encoding.Unicode.GetBytes("SMTP\0")),
            (0x3003001F, Encoding.Unicode.GetBytes(address)));
    }
    var attachment = storage.CreateStorage("__attach_version1.0_#00000000");
    WriteProperties(attachment, 8,
        (0x37050003, BitConverter.GetBytes(1)),
        (0x3707001F, Encoding.Unicode.GetBytes("Anhang.txt\0")),
        (0x37010102, Encoding.UTF8.GetBytes("Anhang Grüße")));
    var inline = storage.CreateStorage("__attach_version1.0_#00000001");
    WriteProperties(inline, 8,
        (0x37050003, BitConverter.GetBytes(1)),
        (0x3707001F, Encoding.Unicode.GetBytes("bild.png\0")),
        (0x3712001F, Encoding.Unicode.GetBytes("bild@example.org\0")),
        (0x37140003, BitConverter.GetBytes(4)),
        (0x37010102, new byte[] { 1, 2, 3 }));
    var nested = storage.CreateStorage("__attach_version1.0_#00000002");
    WriteProperties(nested, 8, (0x37050003, BitConverter.GetBytes(5)));
    WriteMessage(nested.CreateStorage("__substg1.0_3701000D"), false, true);
}

static void WriteProperties(Storage storage, int headerLength, params (uint Tag, byte[] Value)[] properties)
{
    using var stream = storage.CreateStream("__properties_version1.0");
    using var writer = new BinaryWriter(stream);
    writer.Write(new byte[headerLength]);
    foreach (var (tag, value) in properties)
    {
        writer.Write(tag);
        writer.Write(6);
        if ((tag & 0xFFFF) is 0x001F or 0x0102)
        {
            writer.Write(value.Length);
            writer.Write(0);
            using var content = storage.CreateStream($"__substg1.0_{tag:X8}");
            content.Write(value);
        }
        else
        {
            writer.Write(value);
            writer.Write(new byte[8 - value.Length]);
        }
    }
}
