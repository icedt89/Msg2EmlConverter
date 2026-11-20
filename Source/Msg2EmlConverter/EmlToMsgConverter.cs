using System.Text;
using MimeKit;
using MsgKit;
using MsgReader.Outlook;
using Serilog;

namespace Msg2EmlConverter;

public static class EmlToMsgConverter
{
    private static readonly ILogger logger = Log.ForContext(typeof(EmlToMsgConverter));

    public static Storage.Message Convert(MimeMessage eml)
    {
        logger.Information("Converting EML message {Subject}.", eml.Subject);

        using var msg = CreateOutlookMessage(eml);
        CopyHeadersAndMetadata(eml, msg);
        CopyRecipients(eml, msg);
        var attachments = CopyBodyAndAttachments(eml, msg);

        var content = new MemoryStream();
        try
        {
            logger.Debug("Serializing Outlook message with {AttachmentCount} attachments to MSG.", attachments.Count);
            msg.Save(content);
            // MsgKit writes the container; the metadata pass restores properties its attachment API cannot set.
            MsgAttachmentProperties.Preserve(content, attachments);

            content.Position = 0;

            // Reopen the serialized container to return the same reader type accepted by MsgToEmlConverter.
            // The returned message owns the stream and releases it when disposed.
            logger.Debug("Opening the generated MSG container ({ByteCount} bytes).", content.Length);
            var converted = new Storage.Message(content, FileAccess.Read, false);

            logger.Information("EML message {Subject} converted successfully.", eml.Subject);

            return converted;
        }
        catch
        {
            content.Dispose();

            throw;
        }
    }

    private static Email CreateOutlookMessage(MimeMessage eml)
    {
        var from = eml.From.Mailboxes.FirstOrDefault();
        var sender = eml.Sender ?? from;

        // Sender is the actual sender; From is the person represented when mail is sent on someone else's behalf.
        logger.Debug("Mapping actual sender {SenderAddress} and represented From address {FromAddress}.",
            sender?.Address, from?.Address);
        return new Email(
            new Sender(sender?.Address ?? string.Empty, sender?.Name ?? string.Empty),
            new Representing(from?.Address ?? string.Empty, from?.Name ?? string.Empty),
            eml.Subject ?? string.Empty,
            readReceipt: eml.Headers.Contains(HeaderId.DispositionNotificationTo));
    }

    private static void CopyHeadersAndMetadata(MimeMessage eml, Email msg)
    {
        logger.Debug("Copying transport headers, message identifiers and metadata.");

        using var headers = new MemoryStream();

        var options = FormatOptions.Default.Clone();
        // MIME headers use CRLF regardless of the operating system's preferred line ending.
        options.NewLineFormat = NewLineFormat.Dos;
        eml.Headers.WriteTo(options, headers);
        // Store the raw header block so repeated headers are not merged by MsgKit's parser.
        msg.AddProperty(PropertyTags.PR_TRANSPORT_MESSAGE_HEADERS_W, Encoding.UTF8.GetString(headers.ToArray()));
        msg.InternetMessageId = eml.MessageId;
        msg.InReplyToId = eml.InReplyTo;
        msg.InternetReferences = string.Join(" ", eml.References.Select(reference => $"<{reference}>"));
        logger.Debug("Copied {HeaderCount} transport headers. Message-ID: {MessageId}; In-Reply-To: {InReplyTo}; references: {ReferenceCount}.",
            eml.Headers.Count, eml.MessageId, eml.InReplyTo, eml.References.Count);

        if (eml.Headers.Contains(HeaderId.Date) && eml.Date > DateTimeOffset.MinValue)
        {
            // MAPI timestamps are UTC; the transport header retains the original timezone representation.
            msg.SentOn = eml.Date.UtcDateTime;
            msg.ReceivedOn = eml.Date.UtcDateTime;
        }

        msg.Importance = eml.Importance switch
        {
            MimeKit.MessageImportance.High => MsgKit.Enums.MessageImportance.IMPORTANCE_HIGH,
            MimeKit.MessageImportance.Low => MsgKit.Enums.MessageImportance.IMPORTANCE_LOW,
            _ => MsgKit.Enums.MessageImportance.IMPORTANCE_NORMAL
        };

        msg.Priority = eml.Priority switch
        {
            MessagePriority.Urgent => MsgKit.Enums.MessagePriority.PRIO_URGENT,
            MessagePriority.NonUrgent => MsgKit.Enums.MessagePriority.PRIO_NONURGENT,
            _ => MsgKit.Enums.MessagePriority.PRIO_NORMAL
        };
        logger.Debug("Message metadata: sent {SentOn}, importance {Importance}, priority {Priority}.",
            msg.SentOn, msg.Importance, msg.Priority);

        if (eml.Headers[HeaderId.Sensitivity] is { } sensitivity)
        {
            msg.AddProperty(PropertyTags.PR_SENSITIVITY, sensitivity.ToLowerInvariant() switch
            {
                "personal" => 1,
                "private" => 2,
                "company-confidential" => 3,
                _ => 0
            });
        }

        if (eml.Headers.Contains(HeaderId.ReturnReceiptTo))
        {
            msg.AddProperty(PropertyTags.PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED, true);
        }
        logger.Debug("Read receipt requested: {ReadReceipt}; delivery receipt requested: {DeliveryReceipt}.",
            msg.ReadRecipient, eml.Headers.Contains(HeaderId.ReturnReceiptTo));
    }

    private static void CopyRecipients(MimeMessage eml, Email msg)
    {
        logger.Debug("Copying To, Cc, Bcc and Reply-To recipients.");

        // Outlook stores individual recipients; Mailboxes expands MIME groups while the raw headers retain their names.
        foreach (var recipient in eml.To.Mailboxes)
        {
            msg.Recipients.AddTo(recipient.Address, recipient.Name);
        }

        foreach (var recipient in eml.Cc.Mailboxes)
        {
            msg.Recipients.AddCc(recipient.Address, recipient.Name);
        }

        foreach (var recipient in eml.Bcc.Mailboxes)
        {
            msg.Recipients.AddBcc(recipient.Address, recipient.Name);
        }

        foreach (var recipient in eml.ReplyTo.Mailboxes)
        {
            msg.ReplyToRecipients.AddTo(recipient.Address, recipient.Name);
        }
        logger.Debug("Copied {RecipientCount} recipients and {ReplyToCount} reply addresses.",
            msg.Recipients.Count, msg.ReplyToRecipients.Count);
    }

    private static IReadOnlyList<MimeEntity> CopyBodyAndAttachments(MimeMessage eml, Email msg)
    {
        logger.Debug("Copying plain-text and HTML bodies and attachments.");

        var parts = EnumerateBodyParts(eml.Body).ToList();
        var text = parts.OfType<TextPart>().FirstOrDefault(part => !part.IsAttachment && part.IsPlain);
        var html = parts.OfType<TextPart>().FirstOrDefault(part => !part.IsAttachment && part.IsHtml);

        msg.BodyText = text?.Text;
        msg.BodyHtml = html?.Text;

        var attachments = parts.Where(part => !ReferenceEquals(part, text) && !ReferenceEquals(part, html)).ToList();
        logger.Debug("Selected body representations: plain text {HasTextBody}, HTML {HasHtmlBody}; {AttachmentCount} attachments remain.",
            text != null, html != null, attachments.Count);
        foreach (var part in attachments)
        {
            AddAttachment(part, msg);
        }
        return attachments;
    }

    private static IEnumerable<MimeEntity> EnumerateBodyParts(MimeEntity? part)
    {
        if (part is Multipart multipart && !part.IsAttachment)
        {
            logger.Debug("Traversing {MimeType} with {PartCount} child parts; declared start: {Start}.",
                multipart.ContentType.MimeType, multipart.Count, multipart.ContentType.Parameters["start"]);
            foreach (var child in multipart)
            {
                foreach (var leaf in EnumerateBodyParts(child))
                {
                    yield return leaf;
                }
            }
        }
        else if (part != null)
        {
            // An attached multipart is one attachment, not a source of message body text.
            if (part.IsAttachment)
            {
                logger.Debug("Keeping attached MIME entity {MimeType} as one attachment.", part.ContentType.MimeType);
            }
            yield return part;
        }
    }

    private static void AddAttachment(MimeEntity part, Email msg)
    {
        var content = new MemoryStream();
        try
        {
            string fileName;
            switch (part)
            {
                case MimePart file:
                    file.Content?.DecodeTo(content);
                    fileName = string.IsNullOrWhiteSpace(file.FileName) ? "attachment.bin" : file.FileName;

                    break;
                case MessagePart { Message: { } message }:
                    // Keep attached messages as EML files, including their own headers and attachments.
                    logger.Debug("Preserving attached message {Subject} as an EML attachment.", message.Subject);
                    message.WriteTo(content);
                    fileName = Path.ChangeExtension(part.ContentDisposition?.FileName ?? "attached-message.eml", ".eml");

                    break;
                case Multipart multipart:
                    // Keep the MIME headers and boundary so the inverse converter can rebuild this entire subtree.
                    logger.Debug("Serializing a multipart attachment with {PartCount} child parts.", multipart.Count);
                    multipart.WriteTo(content);
                    fileName = part.ContentDisposition?.FileName ?? "attached-part.mime";
                    break;
                default:
                    throw new InvalidDataException("The MIME attachment has no readable content.");
            }

            fileName = GetUniqueAttachmentName(fileName, part.ContentId, msg);

            var inline = MsgAttachmentProperties.IsInlineResource(part);

            logger.Debug("Adding attachment {FileName}: {ByteCount} bytes, MIME type {MimeType}, charset {Charset}, inline {IsInline}.",
                fileName, content.Length, part.ContentType.MimeType, part.ContentType.Charset, inline);

            content.Position = 0;

            // MsgKit owns the attachment stream after it has been added.
            // MsgKit requires a Content-ID for its inline flag; Content-Location-only resources are marked in the metadata pass.
            msg.Attachments.Add(content, fileName,
                isInline: inline && !string.IsNullOrWhiteSpace(part.ContentId),
                contentId: part.ContentId ?? string.Empty);
        }
        catch
        {
            content.Dispose();

            throw;
        }
    }

    private static string GetUniqueAttachmentName(string fileName, string? contentId, Email msg)
    {
        var originalName = Path.GetFileName(fileName);
        var candidate = originalName;
        var suffix = 2;
        while (msg.Attachments.Any(attachment =>
                   string.Equals(attachment.FileName, candidate, StringComparison.OrdinalIgnoreCase)
                   && string.Equals(attachment.ContentId, contentId ?? string.Empty, StringComparison.OrdinalIgnoreCase)))
        {
            candidate = $"{Path.GetFileNameWithoutExtension(originalName)} ({suffix++}){Path.GetExtension(originalName)}";
        }

        if (!string.Equals(originalName, candidate, StringComparison.Ordinal))
        {
            logger.Information("Renaming duplicate attachment {OriginalName} to {FileName}.", originalName, candidate);
        }
        return candidate;
    }
}
