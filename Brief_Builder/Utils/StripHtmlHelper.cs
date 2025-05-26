using System.Text.RegularExpressions;

namespace Brief_Builder.Utils
{
    public class HtmlHelper
    {
        public string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            return Regex.Replace(html, "<[^>]+>", string.Empty);
        }
    }
}


