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

        /// <summary>
        /// Parses decrypted MIME content and extracts the body text/HTML and attachments.
        /// This is the inverse of Build() — it reconstructs the original message parts.
        /// </summary>
        public static MimeExtractResult ExtractBody(string mimeContent)
        {
            var result = new MimeExtractResult();
            if (string.IsNullOrEmpty(mimeContent))
                return result;

            // Check for multipart boundary
            int boundaryIdx = mimeContent.IndexOf("boundary=\"", StringComparison.OrdinalIgnoreCase);
            if (boundaryIdx < 0)
            {
                // Single-part message — decode the base64 body
                return ExtractSinglePart(mimeContent);
            }

            int boundaryStart = boundaryIdx + "boundary=\"".Length;
            int boundaryEnd = mimeContent.IndexOf('"', boundaryStart);
            if (boundaryEnd < 0)
                return result;

            string boundary = mimeContent.Substring(boundaryStart, boundaryEnd - boundaryStart);
            string delimiter = "--" + boundary;
            string terminator = delimiter + "--";

            // Split on boundary
            var parts = new List<string>();
            int pos = mimeContent.IndexOf(delimiter, StringComparison.Ordinal);
            while (pos >= 0)
            {
                int partStart = pos + delimiter.Length;
                // Skip \r\n after delimiter
                if (partStart < mimeContent.Length && mimeContent[partStart] == '\r') partStart++;
                if (partStart < mimeContent.Length && mimeContent[partStart] == '\n') partStart++;

                int nextBoundary = mimeContent.IndexOf(delimiter, partStart, StringComparison.Ordinal);
                if (nextBoundary < 0)
                    break;

                parts.Add(mimeContent.Substring(partStart, nextBoundary - partStart));
                pos = nextBoundary;

                // Check if this is the terminator
                if (nextBoundary + delimiter.Length < mimeContent.Length &&
                    mimeContent.Substring(nextBoundary, terminator.Length) == terminator)
                    break;
            }

            foreach (var part in parts)
            {
                // Split headers from body (blank line separates them)
                int headerEnd = part.IndexOf("\r\n\r\n", StringComparison.Ordinal);
                string nlSeq = "\r\n\r\n";
                if (headerEnd < 0)
                {
                    headerEnd = part.IndexOf("\n\n", StringComparison.Ordinal);
                    nlSeq = "\n\n";
                }
                if (headerEnd < 0)
                    continue;

                string headers = part.Substring(0, headerEnd);
                string body = part.Substring(headerEnd + nlSeq.Length).Trim();

                // Skip protected headers part
                if (headers.IndexOf("protected-headers=", StringComparison.OrdinalIgnoreCase) >= 0)
                    continue;

                // Check Content-Disposition for attachments
                if (headers.IndexOf("Content-Disposition: attachment", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var att = new MimeAttachment();

                    // Extract filename
                    int fnIdx = headers.IndexOf("filename=\"", StringComparison.OrdinalIgnoreCase);
                    if (fnIdx >= 0)
                    {
                        int fnStart = fnIdx + "filename=\"".Length;
                        int fnEnd = headers.IndexOf('"', fnStart);
                        if (fnEnd > fnStart)
                            att.FileName = headers.Substring(fnStart, fnEnd - fnStart);
                    }

                    // Decode base64 body
                    if (headers.IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        try { att.Data = Convert.FromBase64String(body.Replace("\r", "").Replace("\n", "")); }
                        catch { att.Data = Encoding.UTF8.GetBytes(body); }
                    }
                    else
                    {
                        att.Data = Encoding.UTF8.GetBytes(body);
                    }

                    result.Attachments.Add(att);
                    continue;
                }

                // Body part — text/html or text/plain
                bool isHtml = headers.IndexOf("text/html", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isBase64 = headers.IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0;

                string decoded;
                if (isBase64)
                {
                    try
                    {
                        var bytes = Convert.FromBase64String(body.Replace("\r", "").Replace("\n", ""));
                        decoded = Encoding.UTF8.GetString(bytes);
                    }
                    catch
                    {
                        decoded = body;
                    }
                }
                else
                {
                    decoded = body;
                }

                if (isHtml)
                    result.HtmlBody = decoded;
                else if (result.TextBody == null) // take first text/plain
                    result.TextBody = decoded;
            }

            return result;
        }

        private static MimeExtractResult ExtractSinglePart(string mimeContent)
        {
            var result = new MimeExtractResult();

            // Find header/body split
            int headerEnd = mimeContent.IndexOf("\r\n\r\n", StringComparison.Ordinal);
            string nlSeq = "\r\n\r\n";
            if (headerEnd < 0)
            {
                headerEnd = mimeContent.IndexOf("\n\n", StringComparison.Ordinal);
                nlSeq = "\n\n";
            }
            if (headerEnd < 0)
            {
                result.TextBody = mimeContent;
                return result;
            }

            string headers = mimeContent.Substring(0, headerEnd);
            string body = mimeContent.Substring(headerEnd + nlSeq.Length).Trim();

            bool isHtml = headers.IndexOf("text/html", StringComparison.OrdinalIgnoreCase) >= 0;
            bool isBase64 = headers.IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0;

            string decoded;
            if (isBase64)
            {
                try
                {
                    var bytes = Convert.FromBase64String(body.Replace("\r", "").Replace("\n", ""));
                    decoded = Encoding.UTF8.GetString(bytes);
                }
                catch { decoded = body; }
            }
            else
            {
                decoded = body;
            }

            if (isHtml)
                result.HtmlBody = decoded;
            else
                result.TextBody = decoded;

            return result;
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

    /// <summary>
    /// Result of parsing decrypted MIME content back into its component parts.
    /// </summary>
    public class MimeExtractResult
    {
        public string? TextBody { get; set; }
        public string? HtmlBody { get; set; }
        public List<MimeAttachment> Attachments { get; set; } = new List<MimeAttachment>();

        public bool HasContent => TextBody != null || HtmlBody != null;
    }

    public class MimeAttachment
    {
        public string FileName { get; set; } = string.Empty;
        public byte[] Data { get; set; } = Array.Empty<byte>();
    }
}
