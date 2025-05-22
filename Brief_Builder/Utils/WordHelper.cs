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
            IEnumerable<EmailInfo> emails,
            IEnumerable<ImportedFile> importedFiles)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(
                ms, WordprocessingDocumentType.Document, true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body;

                WriteTitle(body, "Brief Builder Report");
                BuildClaimsSection(body, claims);
                BuildEmailsSection(body, emails);
                BuildImportedFilesSection(body, importedFiles);

                mainPart.Document.Save();
            }

            return ms.ToArray();
        }

        private static void WriteTitle(Body body, string text)
        {
            var titlePara = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }),
                new Run(new Text(text)));
            body.Append(titlePara);
        }

        private static void BuildClaimsSection(
            Body body,
            IEnumerable<KeyValuePair<string, string>> claims)
        {
            if (claims == null || !claims.Any()) return;
            body.Append(new Paragraph(new Run(new Text("Claims:"))));
            foreach (var kv in claims)
                body.Append(new Paragraph(new Run(new Text($"- {kv.Key}: {kv.Value}"))));
            body.Append(new Paragraph(new Run(new Text(string.Empty))));
        }

        private static void BuildEmailsSection(
            Body body,
            IEnumerable<EmailInfo> emails)
        {
            if (emails == null || !emails.Any()) return;
            body.Append(new Paragraph(new Run(new Text("Emails:"))));
            foreach (var e in emails)
            {
                AppendLine(body, "Name", e.Name);
                AppendLine(body, "From", e.From);
                AppendLine(body, "To", e.To);
                body.Append(new Paragraph(new Run(new Text(e.Body))));
                body.Append(new Paragraph(new Run(new Text(string.Empty))));
            }
        }

        private static void BuildImportedFilesSection(
            Body body,
            IEnumerable<ImportedFile> files)
        {
            if (files == null || !files.Any()) return;
            body.Append(new Paragraph(new Run(new Text("SharePoint Files:"))));
            foreach (var f in files)
            {
                foreach (var line in f.Text
                                     .Split(new[] { '\r', '\n' },
                                            System.StringSplitOptions.RemoveEmptyEntries))
                {
                    body.Append(new Paragraph(new Run(new Text(line))));
                }
                body.Append(new Paragraph(new Run(new Text(string.Empty))));
            }
        }

        private static void AppendLine(Body body, string label, string value)
        {
            body.Append(new Paragraph(new Run(new Text($"   {label}: {value}"))));
        }
    }
}
