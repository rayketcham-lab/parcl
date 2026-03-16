using System;
using System.Collections.Generic;
using System.Text;

namespace Parcl.Core.Crypto
{
    /// <summary>
    /// Builds a simple MIME message from body + attachments for S/MIME encapsulation.
    /// The resulting bytes are what gets encrypted into the CMS envelope.
    /// </summary>
    public static class MimeBuilder
    {
        public static byte[] Build(string? bodyText, string? bodyHtml, List<MimeAttachment>? attachments)
        {
            bool hasAttachments = attachments != null && attachments.Count > 0;
            bool hasHtml = !string.IsNullOrEmpty(bodyHtml);

            if (!hasAttachments)
            {
                // Simple single-part message
                if (hasHtml)
                    return BuildSinglePart("text/html", Encoding.UTF8.GetBytes(bodyHtml!));
                return BuildSinglePart("text/plain", Encoding.UTF8.GetBytes(bodyText ?? ""));
            }

            // Multipart/mixed: body + attachments
            var boundary = "----=_Parcl_" + Guid.NewGuid().ToString("N").Substring(0, 16);
            var sb = new StringBuilder();
            sb.AppendLine($"MIME-Version: 1.0");
            sb.AppendLine($"Content-Type: multipart/mixed; boundary=\"{boundary}\"");
            sb.AppendLine();

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
