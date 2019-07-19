using System.Text.RegularExpressions;

namespace OneNoteToFiles.Helpers
{
    public static class StringUtils
    {
        private static readonly Regex htmlPattern = new Regex(@"<(.|\n)*?>", RegexOptions.Compiled);

        public static string GetText(string htmlString)
        {
            return htmlPattern.Replace(htmlString, string.Empty);
        }        
    }
}
