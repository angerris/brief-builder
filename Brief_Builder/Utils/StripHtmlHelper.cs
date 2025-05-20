using System.Text.RegularExpressions;

namespace Brief_Builder.Utils
{
    public static class HtmlHelper
    {
        public static string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            return Regex.Replace(html, "<[^>]+>", string.Empty);
        }
    }
}


