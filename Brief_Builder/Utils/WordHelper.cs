using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Brief_Builder.Models;
using DocumentFormat.OpenXml;

namespace Brief_Builder.Utils
{
    public static class WordHelper
    {
        public static byte[] CreateDoc(
            IEnumerable<KeyValuePair<string, string>> claims,
            IEnumerable<EmailInfo> emails)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(
                ms,
                WordprocessingDocumentType.Document,
                true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body;

                var titleParagraph = new Paragraph();
                titleParagraph.ParagraphProperties = new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center });
                var titleRun = new Run();
                titleRun.Append(new DocumentFormat.OpenXml.Wordprocessing.Text("Brief Builder Report"));
                titleParagraph.Append(titleRun);
                body.Append(titleParagraph);

                if (claims != null && claims.Any())
                {
                    body.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Claims:"))));
                    foreach (var kv in claims)
                    {
                        var p = new Paragraph();
                        p.Append(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text($"- {kv.Key}: {kv.Value}")));
                        body.Append(p);
                    }
                    body.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(string.Empty))));
                }

                if (emails != null && emails.Any())
                {
                    body.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Emails:"))));
                    foreach (var e in emails)
                    {
                        var header = new Paragraph();
                        header.Append(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text($"— Email {e.Id}")));
                        body.Append(header);

                        void AppendLine(string label, string value)
                        {
                            var para = new Paragraph();
                            para.Append(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text($"   {label}: {value}")));
                            body.Append(para);
                        }

                        AppendLine("Name", e.Name);
                        AppendLine("From", e.From);
                        AppendLine("To", e.To);
                        AppendLine("Body", e.Body);

                        body.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(string.Empty))));
                    }
                }

                mainPart.Document.Save();
            }

            return ms.ToArray();
        }
    }
}