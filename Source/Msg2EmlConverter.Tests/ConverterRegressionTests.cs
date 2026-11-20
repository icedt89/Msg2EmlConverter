using System.Text;
using Msg2EmlConverter;
using MimeKit;
using MsgReader.Outlook;

internal static class ConverterRegressionTests
{
    public static void Run(string testDirectory, Action<bool, string> check)
    {
        PreserveAttachmentMetadata(testDirectory, check);
        PreserveTextAttachmentCharset(testDirectory, check);
        PreserveContentLocation(testDirectory, check);
        PreserveAttachedMultipart(testDirectory, check);
    }

    private static void PreserveAttachmentMetadata(string testDirectory, Action<bool, string> check)
    {
        using var eml = CreateMessage("Attachment metadata");
        var data = Encoding.ASCII.GetBytes("%PDF-1.4");
        var builder = new BodyBuilder { TextBody = "Body" };
        var attachment = builder.Attachments.Add("document.bin", data, new ContentType("application", "pdf"));
        var created = new DateTimeOffset(2020, 1, 2, 3, 4, 5, TimeSpan.FromHours(2));
        var modified = new DateTimeOffset(2021, 2, 3, 4, 5, 6, TimeSpan.FromHours(-3));
        attachment.ContentDisposition!.CreationDate = created;
        attachment.ContentDisposition.ModificationDate = modified;
        builder.Attachments.Add("other.bin", new byte[] { 4, 5, 6 });
        eml.Body = builder.ToMessageBody();

        using var msg = SaveAndReload(eml, Path.Combine(testDirectory, "metadata.msg"));
        var native = msg.Attachments.OfType<Storage.Attachment>().Single(part => part.FileName == "document.bin");
        check(native.MimeType == "application/pdf", "MSG speichert den ursprünglichen MIME-Typ unabhängig vom Dateinamen");
        check(native.CreationTime == created && native.LastModificationTime == modified, "MSG speichert ursprüngliche Anhang-Zeitstempel als denselben Zeitpunkt");
        var undated = msg.Attachments.OfType<Storage.Attachment>().Single(part => part.FileName == "other.bin");
        check(undated.CreationTime == null && undated.LastModificationTime == null, "Fehlende Anhang-Zeitstempel werden nicht erfunden");

        using var restored = MsgToEmlConverter.Convert(msg);
        var restoredPart = restored.Attachments.OfType<MimePart>().Single(part => part.FileName == "document.bin");
        check(restoredPart.ContentType.MimeType == "application/pdf", "EML erhält den ursprünglichen MIME-Typ nach dem Rückweg");
        check(restoredPart.ContentDisposition?.CreationDate == created && restoredPart.ContentDisposition?.ModificationDate == modified, "EML erhält ursprüngliche Anhang-Zeitstempel nach dem Rückweg");
        using var bytes = new MemoryStream();
        restoredPart.Content!.DecodeTo(bytes);
        check(bytes.ToArray().SequenceEqual(data), "Metadatenkorrektur verändert keine Anhangsdaten");
    }

    private static void PreserveTextAttachmentCharset(string testDirectory, Action<bool, string> check)
    {
        const string text = "Grüße €";
        foreach (var charset in new[] { "windows-1252", "utf-8" })
        {
            using var eml = CreateMessage("Text attachment charset");
            var data = Encoding.GetEncoding(charset).GetBytes(text);
            var builder = new BodyBuilder { TextBody = "Body" };
            builder.Attachments.Add("notes.txt", data, new ContentType("text", "plain") { Charset = charset });
            eml.Body = builder.ToMessageBody();

            using var msg = SaveAndReload(eml, Path.Combine(testDirectory, $"charset-{charset}.msg"));
            var native = msg.Attachments.OfType<Storage.Attachment>().Single();
            check(native.GetMapiProperty("371B") as string == charset, "MSG speichert den Zeichensatz eines Textanhangs");
            using var restored = MsgToEmlConverter.Convert(msg);
            var restoredPart = (TextPart)restored.Attachments.Single();
            check(restoredPart.ContentType.Charset == charset && restoredPart.Text == text, "Textanhang erhält Zeichensatz und Eurozeichen nach dem Rückweg");
            using var bytes = new MemoryStream();
            restoredPart.Content!.DecodeTo(bytes);
            check(bytes.ToArray().SequenceEqual(data), "Zeichensatzkorrektur verändert die Anhangsdaten nicht");

            using var serialized = new MemoryStream();
            restored.WriteTo(serialized);
            serialized.Position = 0;
            using var reloaded = MimeMessage.Load(serialized);
            var reloadedPart = (TextPart)reloaded.Attachments.Single();
            check(reloadedPart.ContentType.Charset == charset && reloadedPart.Text == text, "Zeichensatz und Eurozeichen bleiben auch in der gespeicherten EML erhalten");
        }

        using var unlabelled = CreateMessage("Text attachment without charset");
        var unlabelledBuilder = new BodyBuilder { TextBody = "Body" };
        unlabelledBuilder.Attachments.Add("notes.txt", Encoding.ASCII.GetBytes("ASCII text"), new ContentType("text", "plain"));
        unlabelled.Body = unlabelledBuilder.ToMessageBody();
        using var unlabelledMsg = SaveAndReload(unlabelled, Path.Combine(testDirectory, "no-charset.msg"));
        var nativeUnlabelled = unlabelledMsg.Attachments.OfType<Storage.Attachment>().Single();
        check(nativeUnlabelled.GetMapiProperty("371B") == null, "Ein fehlender Anhang-Zeichensatz wird nicht erfunden");
        using var unlabelledResult = MsgToEmlConverter.Convert(unlabelledMsg);
        var unlabelledPart = (TextPart)unlabelledResult.Attachments.Single();
        check(unlabelledPart.ContentType.Charset == null && unlabelledPart.Text == "ASCII text", "Textanhang ohne Zeichensatz bleibt lesbar und ohne zusätzliche Zeichensatzangabe");
    }

    private static void PreserveContentLocation(string testDirectory, Action<bool, string> check)
    {
        foreach (var (location, contentId, explicitAttachment) in new[]
                 {
                     ("image001.png", (string?)null, false),
                     ("https://example.org/images/logo.png", (string?)null, false),
                     ("images/logo%20small.png", "logo@example.org", false),
                     ("image001.png", (string?)null, true),
                     ("image001.png", "attachment@example.org", true)
                 })
        {
            using var eml = CreateMessage("Content-Location resource");
            var image = new MimePart("image", "png")
            {
                Content = new MimeContent(new MemoryStream(new byte[] { 1, 2, 3 })),
                ContentLocation = new Uri(location, UriKind.RelativeOrAbsolute),
                ContentId = contentId,
                ContentDisposition = new ContentDisposition(explicitAttachment ? ContentDisposition.Attachment : ContentDisposition.Inline)
                {
                    FileName = "image001.png",
                    CreationDate = new DateTimeOffset(2022, 1, 1, 0, 0, 0, TimeSpan.Zero)
                }
            };
            eml.Body = new Multipart("related")
            {
                new TextPart("html") { Text = $"<html><body><img src=\"{location}\"></body></html>" },
                image
            };

            using var msg = SaveAndReload(eml, Path.Combine(testDirectory, "location.msg"));
            var native = msg.Attachments.OfType<Storage.Attachment>().Single();
            check(native.GetMapiProperty("3713") as string == location, "MSG speichert relative und absolute Content-Location");
            check(native.ContentId == contentId, "Content-Location benötigt keine erfundene Content-ID");
            using var restored = MsgToEmlConverter.Convert(msg);
            var restoredImage = restored.BodyParts.OfType<MimePart>().Single(part => part.ContentType.MediaType == "image");
            check(restoredImage.ContentLocation?.OriginalString == location, "EML erhält Content-Location nach dem Rückweg");
            check(restoredImage.IsAttachment == explicitAttachment, "Inline-Ressourcen bleiben inline und explizite Anhänge bleiben Anhänge");
            check(restored.HtmlBody?.TrimEnd('\r', '\n') == eml.HtmlBody?.TrimEnd('\r', '\n'), "HTML-Referenzen bleiben unverändert");
            check(restoredImage.ContentDisposition?.CreationDate == image.ContentDisposition.CreationDate, "Zeitstempel bleiben auch bei Inline-Ressourcen erhalten");
            if (!explicitAttachment)
            {
                check(HasRelatedResource(restored.Body, restoredImage), "Inline-Ressourcen behalten ihren Related-Zusammenhang");
            }
            using var bytes = new MemoryStream();
            restoredImage.Content!.DecodeTo(bytes);
            check(bytes.ToArray().SequenceEqual(new byte[] { 1, 2, 3 }), "Content-Location-Korrektur verändert keine Bilddaten");
        }
    }

    private static void PreserveAttachedMultipart(string testDirectory, Action<bool, string> check)
    {
        foreach (var attachmentFirst in new[] { false, true })
        {
            using var eml = CreateMessage("Attached multipart");
            var bundle = new Multipart("mixed")
            {
                new TextPart("plain") { Text = "TEXT INSIDE ATTACHMENT" },
                new MimePart("application", "zip")
                {
                    Content = new MimeContent(new MemoryStream(new byte[] { 7, 8, 9 })),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment) { FileName = "data.zip" }
                }
            };
            bundle.ContentDisposition = new ContentDisposition(ContentDisposition.Attachment) { FileName = "bundle.mime" };
            var html = new TextPart("html") { Text = "<html><body>Actual message</body></html>" };
            eml.Body = attachmentFirst ? new Multipart("mixed") { bundle, html } : new Multipart("mixed") { html, bundle };

            using var msg = SaveAndReload(eml, Path.Combine(testDirectory, "multipart.msg"));
            check(msg.Attachments.Count == 1, "Multipart-Anhang wird als eine Einheit in MSG gespeichert");
            check(msg.BodyText?.Contains("TEXT INSIDE ATTACHMENT") != true, "Text aus einem Multipart-Anhang wird nicht zum Nachrichtentext");
            using var restored = MsgToEmlConverter.Convert(msg);
            check(restored.TextBody?.Contains("TEXT INSIDE ATTACHMENT") != true && restored.HtmlBody?.TrimEnd('\r', '\n') == html.Text, "Nachrichtentext bleibt unabhängig von der Reihenfolge der Anhänge erhalten");
            var restoredBundle = (restored.Body as Multipart)?.OfType<Multipart>().SingleOrDefault(part => part.IsAttachment);
            check(restoredBundle?.ContentDisposition?.FileName == "bundle.mime", "EML stellt den angehängten Multipart-Knoten wieder her");
            check(restoredBundle?.Count == 2 && restoredBundle[0] is TextPart innerText && innerText.Text.TrimEnd('\r', '\n') == "TEXT INSIDE ATTACHMENT", "Multipart-Anhang behält seine innere Struktur und Textdaten");
            using var nestedBytes = new MemoryStream();
            ((MimePart)restoredBundle![1]).Content!.DecodeTo(nestedBytes);
            check(nestedBytes.ToArray().SequenceEqual(new byte[] { 7, 8, 9 }), "Binärdaten innerhalb des Multipart-Anhangs bleiben erhalten");

            using var serialized = new MemoryStream();
            restored.WriteTo(serialized);
            serialized.Position = 0;
            using var reloaded = MimeMessage.Load(serialized);
            using var secondMsg = EmlToMsgConverter.Convert(reloaded);
            check(secondMsg.Attachments.Count == 1 && secondMsg.BodyText?.Contains("TEXT INSIDE ATTACHMENT") != true, "Multipart-Anhang bleibt auch nach erneutem Speichern und Konvertieren eine Einheit");
        }
    }

    private static bool HasRelatedResource(MimeEntity? entity, MimeEntity resource)
    {
        if (entity is not Multipart multipart)
        {
            return false;
        }
        return (multipart.ContentType.IsMimeType("multipart", "related") && multipart.Contains(resource))
               || multipart.Any(child => HasRelatedResource(child, resource));
    }

    private static Storage.Message SaveAndReload(MimeMessage eml, string path)
    {
        using var converted = EmlToMsgConverter.Convert(eml);
        converted.Save(path);
        return new Storage.Message(path);
    }

    private static MimeMessage CreateMessage(string subject)
    {
        var message = new MimeMessage();
        message.Subject = subject;
        message.From.Add(new MailboxAddress("Author", "author@example.org"));
        message.To.Add(new MailboxAddress("Recipient", "recipient@example.org"));
        return message;
    }
}
