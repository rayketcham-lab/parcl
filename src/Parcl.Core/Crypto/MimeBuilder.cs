using System;
using System.Collections.Generic;
using System.Text;

namespace Parcl.Core.Crypto
{
    /// <summary>
    /// RFC 7508 protected headers to include inside the encrypted MIME envelope.
    /// These headers are hidden from the outer (cleartext) message.
    /// </summary>
    public class ProtectedHeaders
    {
        public string? Subject { get; set; }
        public string? From { get; set; }
        public string? To { get; set; }
        public string? Date { get; set; }
    }

    /// <summary>
    /// Builds a simple MIME message from body + attachments for S/MIME encapsulation.
    /// The resulting bytes are what gets encrypted into the CMS envelope.
    /// </summary>
    public static class MimeBuilder
    {
        public static byte[] Build(string? bodyText, string? bodyHtml,
            List<MimeAttachment>? attachments, ProtectedHeaders? headers = null)
        {
            bool hasAttachments = attachments != null && attachments.Count > 0;
            bool hasHtml = !string.IsNullOrEmpty(bodyHtml);
            bool hasProtectedHeaders = headers != null;

            // When protected headers are present, always use multipart/mixed
            // so the headers part is the first MIME part (RFC 7508).
            if (!hasAttachments && !hasProtectedHeaders)
            {
                // Simple single-part message
                if (hasHtml)
                    return BuildSinglePart("text/html", Encoding.UTF8.GetBytes(bodyHtml!));
                return BuildSinglePart("text/plain", Encoding.UTF8.GetBytes(bodyText ?? ""));
            }

            // Multipart/mixed: [protected headers] + body + attachments
            var boundary = "----=_Parcl_" + Guid.NewGuid().ToString("N").Substring(0, 16);
            var sb = new StringBuilder();
            sb.AppendLine($"MIME-Version: 1.0");
            sb.AppendLine($"Content-Type: multipart/mixed; boundary=\"{boundary}\"");
            sb.AppendLine();

            // Protected headers part (RFC 7508) — must be first
            if (hasProtectedHeaders)
            {
                sb.AppendLine($"--{boundary}");
                sb.AppendLine("Content-Type: text/rfc822-headers; protected-headers=\"v1\"");

                if (!string.IsNullOrEmpty(headers!.Subject))
                    sb.AppendLine($"Subject: {SanitizeHeaderValue(headers.Subject!)}");
                if (!string.IsNullOrEmpty(headers.From))
                    sb.AppendLine($"From: {SanitizeHeaderValue(headers.From!)}");
                if (!string.IsNullOrEmpty(headers.To))
                    sb.AppendLine($"To: {SanitizeHeaderValue(headers.To!)}");
                if (!string.IsNullOrEmpty(headers.Date))
                    sb.AppendLine($"Date: {SanitizeHeaderValue(headers.Date!)}");

                sb.AppendLine();
            }

            // Body part
            sb.AppendLine($"--{boundary}");
            if (hasHtml)
            {
                sb.AppendLine("Content-Type: text/html; charset=utf-8");
                sb.AppendLine("Content-Transfer-Encoding: base64");
                sb.AppendLine();
                sb.AppendLine(WrapBase64(Convert.ToBase64String(Encoding.UTF8.GetBytes(bodyHtml!))));
            }
            else
            {
                sb.AppendLine("Content-Type: text/plain; charset=utf-8");
                sb.AppendLine("Content-Transfer-Encoding: base64");
                sb.AppendLine();
                sb.AppendLine(WrapBase64(Convert.ToBase64String(Encoding.UTF8.GetBytes(bodyText ?? ""))));
            }

            // Attachment parts
            if (hasAttachments)
            {
                foreach (var att in attachments!)
                {
                    var safeName = SanitizeFilename(att.FileName);
                    sb.AppendLine($"--{boundary}");
                    sb.AppendLine($"Content-Type: application/octet-stream; name=\"{safeName}\"");
                    sb.AppendLine("Content-Transfer-Encoding: base64");
                    sb.AppendLine($"Content-Disposition: attachment; filename=\"{safeName}\"");
                    sb.AppendLine();
                    sb.AppendLine(WrapBase64(Convert.ToBase64String(att.Data)));
                }
            }

            sb.AppendLine($"--{boundary}--");

            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static byte[] BuildSinglePart(string contentType, byte[] body)
        {
            var sb = new StringBuilder();
            sb.AppendLine("MIME-Version: 1.0");
            sb.AppendLine($"Content-Type: {contentType}; charset=utf-8");
            sb.AppendLine("Content-Transfer-Encoding: base64");
            sb.AppendLine();
            sb.AppendLine(WrapBase64(Convert.ToBase64String(body)));
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        /// <summary>
        /// Strips CR and LF from header values to prevent MIME header injection.
        /// </summary>
        public static string SanitizeHeaderValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            var sb = new StringBuilder(value.Length);
            foreach (char c in value)
            {
                if (c == '\r' || c == '\n' || c == '\0')
                    continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Parses decrypted MIME content for RFC 7508 protected headers.
        /// Returns a ProtectedHeaders instance if found, null otherwise.
        /// </summary>
        public static ProtectedHeaders? ExtractProtectedHeaders(string mimeContent)
        {
            if (string.IsNullOrEmpty(mimeContent))
                return null;

            // Look for the protected-headers="v1" marker
            int markerIndex = mimeContent.IndexOf("protected-headers=\"v1\"",
                StringComparison.OrdinalIgnoreCase);
            if (markerIndex < 0)
                return null;

            // Find the headers block: starts after the Content-Type line,
            // ends at the first blank line
            int lineStart = mimeContent.LastIndexOf('\n', markerIndex);
            if (lineStart < 0) lineStart = 0;
            else lineStart++;

            // Skip past the Content-Type line itself
            int afterContentType = mimeContent.IndexOf('\n', markerIndex);
            if (afterContentType < 0)
                return null;
            afterContentType++;

            // Find end of header block (blank line)
            int blockEnd = mimeContent.IndexOf("\r\n\r\n", afterContentType, StringComparison.Ordinal);
            if (blockEnd < 0)
                blockEnd = mimeContent.IndexOf("\n\n", afterContentType, StringComparison.Ordinal);
            if (blockEnd < 0)
                return null;

            string headerBlock = mimeContent.Substring(afterContentType, blockEnd - afterContentType);
            var result = new ProtectedHeaders();
            bool found = false;

            foreach (string rawLine in headerBlock.Split('\n'))
            {
                string line = rawLine.TrimEnd('\r');
                if (line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
                {
                    result.Subject = line.Substring("Subject:".Length).Trim();
                    found = true;
                }
                else if (line.StartsWith("From:", StringComparison.OrdinalIgnoreCase))
                {
                    result.From = line.Substring("From:".Length).Trim();
                    found = true;
                }
                else if (line.StartsWith("To:", StringComparison.OrdinalIgnoreCase))
                {
                    result.To = line.Substring("To:".Length).Trim();
                    found = true;
                }
                else if (line.StartsWith("Date:", StringComparison.OrdinalIgnoreCase))
                {
                    result.Date = line.Substring("Date:".Length).Trim();
                    found = true;
                }
            }

            return found ? result : null;
        }

        /// <summary>
        /// Strips CR, LF, NUL, and double-quote characters from filenames
        /// to prevent MIME header injection.
        /// </summary>
        public static string SanitizeFilename(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                return filename;

            var sb = new StringBuilder(filename.Length);
            foreach (char c in filename)
            {
                if (c == '\r' || c == '\n' || c == '\0' || c == '"')
                    continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        private static string WrapBase64(string b64)
        {
            var sb = new StringBuilder(b64.Length + b64.Length / 76);
            for (int i = 0; i < b64.Length; i += 76)
            {
                sb.AppendLine(b64.Substring(i, Math.Min(76, b64.Length - i)));
            }
            return sb.ToString().TrimEnd();
        }
    }

    public class MimeAttachment
    {
        public string FileName { get; set; } = string.Empty;
        public byte[] Data { get; set; } = Array.Empty<byte>();
    }
}
