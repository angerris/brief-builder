using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Brief_Builder.Models;
using DocumentFormat.OpenXml;
using System.Linq;
using System.Text.RegularExpressions;

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
                ms,
                WordprocessingDocumentType.Document,
                true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body;

                WriteTitle(body, "Brief Builder Report");
                body.Append(new Paragraph(new Run(new Text(string.Empty))));
                BuildClaimsSection(body, claims);
                BuildEmailsSection(body, emails);
                AppendImportedDocs(mainPart, body, importedFiles);

                mainPart.Document.Save();
            }
            return ms.ToArray();
        }

        private static void WriteTitle(Body body, string text)
        {
            var para = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }),
                new Run(
                    new RunProperties(
                        new Bold(),
                        new FontSize { Val = "32" }
                    ),
                    new Text(text)
                ));
            body.Append(para);
        }

        private static void WriteSubtitle(Body body, string text)
        {
            var para = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Left }),
                new Run(
                    new RunProperties(
                        new Bold(),
                        new FontSize { Val = "28" }
                    ),
                    new Text(text)
                ));
            body.Append(para);
        }

        private static void BuildClaimsSection(
            Body body,
            IEnumerable<KeyValuePair<string, string>> claims)
        {
            if (claims == null || !claims.Any()) return;

            WriteSubtitle(body, "Claims");
            foreach (var kv in claims)
            {
                body.Append(new Paragraph(
                    new Run(new Text($"- {kv.Key}: {kv.Value}"))));
            }
            body.Append(new Paragraph(new Run(new Text(string.Empty))));
            body.Append(new Paragraph(new Run(new Text(string.Empty))));
        }

        private static void BuildEmailsSection(
            Body body,
            IEnumerable<EmailInfo> emails)
        {
            if (emails == null || !emails.Any()) return;

            WriteSubtitle(body, "Emails");
            foreach (var e in emails)
            {
                body.Append(new Paragraph(
                    new Run(new Text($"Name: {e.Name}"))));
                body.Append(new Paragraph(
                    new Run(new Text($"From: {e.From}"))));
                body.Append(new Paragraph(
                    new Run(new Text($"To:   {e.To}"))));
                body.Append(new Paragraph(
                    new Run(new Text(e.Body))));
                body.Append(new Paragraph(new Run(new Text(string.Empty))));
            }
        }

        private static void AppendImportedDocs(
            MainDocumentPart mainPart,
            Body body,
            IEnumerable<ImportedFile> importedFiles)
            {
                if (importedFiles == null) return;

                WriteSubtitle(body, "SharePoint Files");

                int chunkId = 0;
                foreach (var file in importedFiles)
                {
                    body.Append(new Paragraph(
                        new Run(new Break { Type = BreakValues.Page })));

                    var displayName = Regex.Replace(file.Name, @"\.[^.]+$", "");

                    var fileNamePara = new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Left }),
                        new Run(new Text(displayName)));

                    body.Append(fileNamePara);

                    var partId = $"altChunkId{++chunkId}";
                    var chunk = mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML,
                        partId);
                    using var stream = new MemoryStream(file.Content);
                    chunk.FeedData(stream);
                    body.Append(new AltChunk { Id = partId });
                }

                body.Append(new Paragraph(
                    new Run(new Break { Type = BreakValues.Page })));
            }
    }
}
