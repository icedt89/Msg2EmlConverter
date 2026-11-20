using System.Text;
using MimeKit;
using MimeKit.Utils;
using MsgReader.Outlook;
using Serilog;

namespace Msg2EmlConverter;

public static class MsgToEmlConverter
{
    private static readonly ILogger logger = Log.ForContext(typeof(MsgToEmlConverter));

    public static MimeMessage Convert(Storage.Message msg)
    {
        logger.Information("Converting MSG message {Subject}.", msg.Subject);

        var eml = new MimeMessage();

        // Remove constructor defaults so missing headers can be filled from the original MAPI values.
        eml.Headers.Clear();

        logger.Debug("Copying transport headers and filling missing sender and recipient addresses.");

        CopyTransportHeaders(msg, eml);
        AddMissingSenderAddresses(msg, eml);
        AddMissingRecipients(msg, eml);
        logger.Debug("Copying subject, date, message identifiers and metadata.");

        SetMissingSubjectAndDate(msg, eml);
        SetMissingMessageIdentifiers(msg, eml);
        SetMissingImportance(msg, eml);
        SetMissingSensitivity(msg, eml);
        AddReceiptRequestHeaders(msg, eml);
        AddCategoryHeader(msg, eml);
        AddFollowUpHeaders(msg, eml);

        logger.Debug("Resolved addresses: From {From}, Sender {Sender}; To {ToCount}, Cc {CcCount}, Bcc {BccCount}, Reply-To {ReplyToCount}.",
            eml.From, eml.Sender, eml.To.Count, eml.Cc.Count, eml.Bcc.Count, eml.ReplyTo.Count);
        logger.Debug("Resolved message metadata: Date {Date}, Message-ID {MessageId}, importance {Importance}.",
            eml.Headers[HeaderId.Date], eml.MessageId, eml.Importance);
        eml.Body = CreateMessageBody(msg);

        logger.Information("MSG message {Subject} converted successfully.", msg.Subject);

        return eml;
    }

    private static void CopyTransportHeaders(Storage.Message msg, MimeMessage eml)
    {
        // Preserve Reply-To and other existing SMTP headers.
        if (string.IsNullOrWhiteSpace(msg.TransportMessageHeaders))
        {
            logger.Debug("No transport headers found; reconstructing available addresses and metadata from MAPI.");
            return;
        }

        using var headerStream = new MemoryStream(Encoding.UTF8.GetBytes(msg.TransportMessageHeaders + "\r\n\r\n"));
        var parsedHeaders = HeaderList.Load(headerStream);

        // Exclude MIME structure headers because MimeKit generates new boundaries when building the body.
        var mimeStructureHeaders = new HashSet<HeaderId>
        {
            HeaderId.ContentType,
            HeaderId.ContentTransferEncoding,
            HeaderId.ContentDisposition,
            HeaderId.MimeVersion
        };

        foreach (var header in parsedHeaders)
        {
            if (!mimeStructureHeaders.Contains(header.Id))
            {
                eml.Headers.Add(header);
            }
        }
        logger.Debug("Copied {HeaderCount} transport headers; excluded {ExcludedCount} MIME structure headers.",
            eml.Headers.Count, parsedHeaders.Count - eml.Headers.Count);
    }

    private static void AddMissingSenderAddresses(Storage.Message msg, MimeMessage eml)
    {
        // Fall back to MAPI properties for locally saved MSG files without transport headers.
        if (eml.From.Count == 0)
        {
            if (msg.SenderRepresenting is { } representing && !string.IsNullOrWhiteSpace(representing.Email))
            {
                eml.From.Add(CreateSafeMailbox(representing.DisplayName, representing.Email));
            }
            else if (msg.Sender is { } sender)
            {
                eml.From.Add(CreateSafeMailbox(sender.DisplayName, sender.Email));
            }
        }

        // Preserve the distinction between Sender and From for messages sent on behalf of another user.
        // An existing From header must not prevent a missing Sender header from being reconstructed.
        if (eml.Sender != null
            || msg is not { Sender: { } actualSender, SenderRepresenting: { } representedSender }
            || string.IsNullOrWhiteSpace(representedSender.Email))
        {
            return;
        }

        if (!string.Equals(actualSender.Email, representedSender.Email, StringComparison.OrdinalIgnoreCase))
        {
            eml.Sender = CreateSafeMailbox(actualSender.DisplayName, actualSender.Email);
        }
    }

    private static void AddMissingRecipients(Storage.Message msg, MimeMessage eml)
    {
        // Each list has its own fallback; transport headers often omit Bcc even when To is present.
        foreach (var type in new[] { RecipientType.To, RecipientType.Cc, RecipientType.Bcc })
        {
            var addresses = type switch
            {
                RecipientType.Cc => eml.Cc,
                RecipientType.Bcc => eml.Bcc,
                _ => eml.To
            };

            if (addresses.Count > 0)
            {
                continue;
            }

            foreach (var recipient in msg.Recipients.Where(recipient => recipient.Type == type))
            {
                addresses.Add(CreateSafeMailbox(recipient.DisplayName, recipient.Email));
            }
        }
    }

    private static void SetMissingSubjectAndDate(Storage.Message msg, MimeMessage eml)
    {
        if (string.IsNullOrEmpty(eml.Subject) && !string.IsNullOrEmpty(msg.Subject))
        {
            eml.Subject = msg.Subject;
        }

        if (eml.Headers.Contains(HeaderId.Date))
        {
            return;
        }

        if (msg.SentOn.HasValue)
        {
            eml.Date = msg.SentOn.Value;
        }
        else if (msg.ReceivedOn.HasValue)
        {
            eml.Date = msg.ReceivedOn.Value;
        }
        else if (msg.CreationTime.HasValue)
        {
            eml.Date = msg.CreationTime.Value;
        }
    }

    private static void SetMissingMessageIdentifiers(Storage.Message msg, MimeMessage eml)
    {
        if (string.IsNullOrEmpty(eml.MessageId) && !string.IsNullOrWhiteSpace(msg.Id))
        {
            eml.MessageId = msg.Id.Trim('<', '>');
        }

        // MAPI: PidTagInReplyToId (1042), PidTagInternetReferences (1039).
        if (string.IsNullOrEmpty(eml.InReplyTo)
            && msg.GetMapiProperty("1042") is string inReplyTo
            && !string.IsNullOrWhiteSpace(inReplyTo))
        {
            eml.InReplyTo = inReplyTo.Trim('<', '>');
        }

        if (eml.References.Count != 0 || msg.GetMapiProperty("1039") is not string references)
        {
            return;
        }

        foreach (var reference in MimeUtils.EnumerateReferences(references))
        {
            eml.References.Add(reference);
        }
    }

    private static void SetMissingImportance(Storage.Message msg, MimeMessage eml)
    {
        if (eml.Headers.Contains(HeaderId.Importance))
        {
            return;
        }

        switch (msg.Importance)
        {
            case MsgReader.Outlook.MessageImportance.High:
                eml.Importance = MimeKit.MessageImportance.High;
                eml.XPriority = XMessagePriority.Highest;

                break;
            case MsgReader.Outlook.MessageImportance.Low:
                eml.Importance = MimeKit.MessageImportance.Low;
                eml.XPriority = XMessagePriority.Lowest;

                break;
        }
    }

    private static void SetMissingSensitivity(Storage.Message msg, MimeMessage eml)
    {
        // MAPI: PidTagSensitivity (0036).
        if (!eml.Headers.Contains(HeaderId.Sensitivity)
            && msg.GetMapiProperty("0036") is int sensitivity
            && sensitivity != 0)
        {
            eml.Headers.Add(HeaderId.Sensitivity, sensitivity switch
            {
                1 => "Personal",
                2 => "Private",
                3 => "Company-Confidential",
                _ => "Normal"
            });
        }
    }

    private static void AddReceiptRequestHeaders(Storage.Message msg, MimeMessage eml)
    {
        // MAPI: PidTagReadReceiptRequested (0029), PidTagOriginatorDeliveryReportRequested (0023).
        if (msg.GetMapiProperty("0029") is true && !eml.Headers.Contains(HeaderId.DispositionNotificationTo))
        {
            var notifyTarget = eml.ReplyTo.FirstOrDefault() ?? eml.From.FirstOrDefault();
            if (notifyTarget != null)
            {
                eml.Headers.Add(HeaderId.DispositionNotificationTo, notifyTarget.ToString());
            }
        }

        if (msg.GetMapiProperty("0023") is not true || eml.Headers.Contains(HeaderId.ReturnReceiptTo))
        {
            return;
        }

        var receiptTarget = eml.From.FirstOrDefault();
        if (receiptTarget != null)
        {
            eml.Headers.Add(HeaderId.ReturnReceiptTo, receiptTarget.ToString());
        }
    }

    private static void AddCategoryHeader(Storage.Message msg, MimeMessage eml)
    {
        if (msg.Categories is { Count: > 0 }
            && !eml.Headers.Contains(HeaderId.Keywords))
        {
            eml.Headers.Add(HeaderId.Keywords, string.Join(", ", msg.Categories));
        }
    }

    private static void AddFollowUpHeaders(Storage.Message msg, MimeMessage eml)
    {
        if (msg.Flag == null)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(msg.Flag.Request)
            && !eml.Headers.Contains("X-Message-Flag"))
        {
            eml.Headers.Add("X-Message-Flag", msg.Flag.Request);
        }

        if (msg.Task?.DueDate is DateTimeOffset dueDate
            && !eml.Headers.Contains("Reply-By"))
        {
            eml.Headers.Add("Reply-By", DateUtils.FormatDate(dueDate));
        }
    }

    private static MimeEntity CreateMessageBody(Storage.Message msg)
    {
        logger.Debug("Building plain-text and HTML bodies with {AttachmentCount} attachments.", msg.Attachments.Count);

        var builder = new BodyBuilder
        {
            TextBody = msg.BodyText,
            HtmlBody = msg.BodyHtml
        };
        logger.Debug("Body sizes: plain text {TextLength} characters; HTML {HtmlLength} characters.",
            builder.TextBody?.Length ?? 0, builder.HtmlBody?.Length ?? 0);

        AddAttachments(msg, builder);

        return builder.ToMessageBody();
    }

    private static void AddAttachments(Storage.Message msg, BodyBuilder builder)
    {
        foreach (var attachment in msg.Attachments)
        {
            switch (attachment)
            {
                case Storage.Attachment { Data: not null } fileAtt:
                    AddFileAttachment(fileAtt, builder);

                    break;
                case Storage.Message embeddedMsg:
                    AddEmbeddedMessageAttachment(embeddedMsg, builder);

                    break;
            }
        }
    }

    private static void AddFileAttachment(Storage.Attachment fileAtt, BodyBuilder builder)
    {
        var fileName = !string.IsNullOrWhiteSpace(fileAtt.FileName)
            ? fileAtt.FileName
            : "anhang.bin";

        var contentType = !string.IsNullOrWhiteSpace(fileAtt.MimeType)
                          && ContentType.TryParse(fileAtt.MimeType, out var parsedCt)
            ? parsedCt
            : ContentType.Parse(MimeTypes.GetMimeType(fileName));

        // MAPI: PidTagTextAttachmentCharset (371B).
        if (contentType.MediaType.Equals("text", StringComparison.OrdinalIgnoreCase)
            && fileAtt.GetMapiProperty("371B") is string charset
            && !string.IsNullOrWhiteSpace(charset))
        {
            contentType.Charset = charset;
            logger.Debug("Restoring text attachment charset {Charset} for {FileName}.", charset, fileName);
        }

        if (contentType.MediaType.Equals("multipart", StringComparison.OrdinalIgnoreCase)
            && TryAddMultipartAttachment(fileAtt, fileName, builder))
        {
            return;
        }

        // MAPI: PidTagAttachContentLocation (3713), PidTagAttachFlags (3714).
        Uri.TryCreate(fileAtt.GetMapiProperty("3713") as string, UriKind.RelativeOrAbsolute, out var contentLocation);
        var inline = fileAtt.GetMapiProperty("3714") is int flags
            ? (flags & 4) != 0
            : fileAtt.IsInline || (fileAtt.Hidden && contentLocation != null);

        MimeEntity addedPart;
        if (inline && (!string.IsNullOrEmpty(fileAtt.ContentId) || contentLocation != null))
        {
            logger.Debug("Adding inline resource {FileName}: {ByteCount} bytes, Content-ID {ContentId}, Content-Location {ContentLocation}.",
                fileName, fileAtt.Data.Length, fileAtt.ContentId, contentLocation);

            var linked = builder.LinkedResources.Add(fileName, fileAtt.Data, contentType);
            if (!string.IsNullOrEmpty(fileAtt.ContentId))
            {
                linked.ContentId = fileAtt.ContentId.Trim('<', '>');
            }
            addedPart = linked;
        }
        else
        {
            logger.Debug("Adding file attachment {FileName}: {ByteCount} bytes, MIME type {MimeType}, charset {Charset}.",
                fileName, fileAtt.Data.Length, contentType.MimeType, contentType.Charset);

            addedPart = builder.Attachments.Add(fileName, fileAtt.Data, contentType);
        }

        if (contentLocation != null)
        {
            addedPart.ContentLocation = contentLocation;
        }
        CopyAttachmentDates(fileAtt, addedPart);
    }

    private static bool TryAddMultipartAttachment(Storage.Attachment fileAtt, string fileName, BodyBuilder builder)
    {
        try
        {
            using var content = new MemoryStream(fileAtt.Data);
            var entity = MimeEntity.Load(content);
            if (entity is Multipart multipart)
            {
                // Restore the serialized subtree instead of exposing its children as body text or separate attachments.
                multipart.ContentDisposition ??= new ContentDisposition(ContentDisposition.Attachment);
                multipart.ContentDisposition.Disposition = ContentDisposition.Attachment;
                multipart.ContentDisposition.FileName = fileName;
                CopyAttachmentDates(fileAtt, multipart);
                builder.Attachments.Add(multipart);
                logger.Debug("Restored multipart attachment {FileName} with {PartCount} child parts.", fileName, multipart.Count);
                return true;
            }
            entity.Dispose();
        }
        catch (FormatException ex)
        {
            logger.Warning(ex, "Multipart attachment {FileName} could not be parsed; preserving its raw data.", fileName);
        }
        return false;
    }

    private static void CopyAttachmentDates(Storage.Attachment fileAtt, MimeEntity part)
    {
        if (part.ContentDisposition == null)
        {
            return;
        }
        if (fileAtt.CreationTime.HasValue)
        {
            part.ContentDisposition.CreationDate = fileAtt.CreationTime.Value;
        }
        if (fileAtt.LastModificationTime.HasValue)
        {
            part.ContentDisposition.ModificationDate = fileAtt.LastModificationTime.Value;
        }
        if (fileAtt.CreationTime.HasValue || fileAtt.LastModificationTime.HasValue)
        {
            logger.Debug("Restored attachment timestamps for {FileName}: created {CreationTime}, modified {ModificationTime}.",
                fileAtt.FileName, fileAtt.CreationTime, fileAtt.LastModificationTime);
        }
    }

    private static void AddEmbeddedMessageAttachment(Storage.Message embeddedMsg, BodyBuilder builder)
    {
        logger.Debug("Converting embedded MSG attachment {Subject}.", embeddedMsg.Subject);

        // Recursively convert attached MSG files and embed them as message/rfc822 parts.
        var nestedMime = Convert(embeddedMsg);
        var messagePart = new MessagePart { Message = nestedMime };

        var subjectName = string.IsNullOrWhiteSpace(embeddedMsg.Subject)
            ? "weitergeleitete_nachricht.eml"
            : $"{SanitizeFileName(embeddedMsg.Subject)}.eml";

        messagePart.ContentDisposition = new ContentDisposition(ContentDisposition.Attachment)
        {
            FileName = subjectName
        };

        builder.Attachments.Add(messagePart);
    }

    private static MailboxAddress CreateSafeMailbox(string? displayName, string? email)
    {
        if (!string.IsNullOrWhiteSpace(email)
            && email.Contains('@')
            && MailboxAddress.TryParse(email, out var parsed))
        {
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                parsed.Name = displayName;
            }

            return parsed;
        }

        // Use a placeholder for internal Exchange X.500 paths (/O=EXCHANGELABS/...) without an SMTP address.
        var safeName = !string.IsNullOrWhiteSpace(displayName)
            ? displayName
            : email ?? "Unbekannt";

        logger.Warning("No usable SMTP address found for {DisplayName}; using internal@exchange.local.", safeName);
        return new MailboxAddress(safeName, "internal@exchange.local");
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var character in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(character, '_');
        }
        return name;
    }
}
