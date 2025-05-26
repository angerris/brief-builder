using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Brief_Builder.Utils
{
    public class DocxHelper
    {
        public string ExtractText(byte[] docxBytes)
        {
            using var ms = new MemoryStream(docxBytes);
            using var word = WordprocessingDocument.Open(ms, false);
            var body = word.MainDocumentPart?.Document.Body;
            if (body == null) return string.Empty;

            return string.Join(" ",
                body
                .Descendants<Text>()
                .Select(t => t.Text)
                .Where(t => !string.IsNullOrEmpty(t)));
        }
    }
}
