using System;
using System.Collections.Generic;
using System.Text;
using Parcl.Core.Crypto;
using Xunit;

namespace Parcl.Core.Tests
{
    public class MimeBuilderTests
    {
        [Fact]
        public void RoundTrip_PlainText_NoAttachments()
        {
            var body = "Hello, this is a test message.";
            var mime = MimeBuilder.Build(body, null, null);
            var mimeText = Encoding.UTF8.GetString(mime);
            var result = MimeBuilder.ExtractBody(mimeText);

            Assert.True(result.HasContent);
            Assert.Equal(body, result.TextBody);
            Assert.Null(result.HtmlBody);
            Assert.Empty(result.Attachments);
        }

        [Fact]
        public void RoundTrip_HtmlBody_NoAttachments()
        {
            var html = "<html><body><p>Hello <b>World</b></p></body></html>";
            var mime = MimeBuilder.Build(null, html, null);
            var mimeText = Encoding.UTF8.GetString(mime);
            var result = MimeBuilder.ExtractBody(mimeText);

            Assert.True(result.HasContent);
            Assert.Equal(html, result.HtmlBody);
            Assert.Empty(result.Attachments);
        }

        [Fact]
        public void RoundTrip_HtmlBody_WithProtectedHeaders()
        {
            var html = "<html><body><p>Encrypted content</p></body></html>";
            var headers = new ProtectedHeaders
            {
                Subject = "Secret Subject",
                From = "alice@example.com",
                To = "bob@example.com",
                Date = "Mon, 16 Mar 2026 20:00:00 GMT"
            };
            var mime = MimeBuilder.Build(null, html, null, headers);
            var mimeText = Encoding.UTF8.GetString(mime);

            var extractedHeaders = MimeBuilder.ExtractProtectedHeaders(mimeText);
            Assert.NotNull(extractedHeaders);
            Assert.Equal("Secret Subject", extractedHeaders!.Subject);
            Assert.Equal("alice@example.com", extractedHeaders.From);
            Assert.Equal("bob@example.com", extractedHeaders.To);

            var result = MimeBuilder.ExtractBody(mimeText);
            Assert.True(result.HasContent);
            Assert.Equal(html, result.HtmlBody);
            Assert.Empty(result.Attachments);
        }

        [Fact]
        public void RoundTrip_WithAttachments()
        {
            var body = "Message with attachment";
            var attachments = new List<MimeAttachment>
            {
                new MimeAttachment
                {
                    FileName = "test.txt",
                    Data = Encoding.UTF8.GetBytes("attachment content")
                }
            };
            var mime = MimeBuilder.Build(body, null, attachments);
            var mimeText = Encoding.UTF8.GetString(mime);
            var result = MimeBuilder.ExtractBody(mimeText);

            Assert.True(result.HasContent);
            Assert.Equal(body, result.TextBody);
            Assert.Single(result.Attachments);
            Assert.Equal("test.txt", result.Attachments[0].FileName);
            Assert.Equal("attachment content", Encoding.UTF8.GetString(result.Attachments[0].Data));
        }

        [Fact]
        public void RoundTrip_HtmlBody_WithAttachmentsAndHeaders()
        {
            var html = "<p>Full round trip</p>";
            var headers = new ProtectedHeaders { Subject = "Test Subject" };
            var attachments = new List<MimeAttachment>
            {
                new MimeAttachment
                {
                    FileName = "doc.pdf",
                    Data = new byte[] { 0x25, 0x50, 0x44, 0x46 } // %PDF
                },
                new MimeAttachment
                {
                    FileName = "image.png",
                    Data = new byte[] { 0x89, 0x50, 0x4E, 0x47 } // PNG header
                }
            };
            var mime = MimeBuilder.Build(null, html, attachments, headers);
            var mimeText = Encoding.UTF8.GetString(mime);
            var result = MimeBuilder.ExtractBody(mimeText);

            Assert.True(result.HasContent);
            Assert.Equal(html, result.HtmlBody);
            Assert.Equal(2, result.Attachments.Count);
            Assert.Equal("doc.pdf", result.Attachments[0].FileName);
            Assert.Equal("image.png", result.Attachments[1].FileName);
            Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46 }, result.Attachments[0].Data);
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, result.Attachments[1].Data);
        }

        [Fact]
        public void ExtractBody_EmptyContent_ReturnsNoContent()
        {
            var result = MimeBuilder.ExtractBody("");
            Assert.False(result.HasContent);

            var resultNull = MimeBuilder.ExtractBody(null!);
            Assert.False(resultNull.HasContent);
        }

        [Fact]
        public void ExtractProtectedHeaders_NoHeaders_ReturnsNull()
        {
            var mime = MimeBuilder.Build("plain text", null, null);
            var mimeText = Encoding.UTF8.GetString(mime);
            var result = MimeBuilder.ExtractProtectedHeaders(mimeText);
            Assert.Null(result);
        }

        [Fact]
        public void SanitizeHeaderValue_StripsInjection()
        {
            Assert.Equal("SafeValue", MimeBuilder.SanitizeHeaderValue("Safe\r\nValue"));
            Assert.Equal("NoNull", MimeBuilder.SanitizeHeaderValue("No\0Null"));
        }

        [Fact]
        public void SanitizeFilename_StripsQuotesAndNewlines()
        {
            Assert.Equal("file.txt", MimeBuilder.SanitizeFilename("file.txt"));
            Assert.Equal("bad.txt", MimeBuilder.SanitizeFilename("bad\".txt"));
            Assert.Equal("bad.txt", MimeBuilder.SanitizeFilename("bad\r\n.txt"));
        }
    }
}
