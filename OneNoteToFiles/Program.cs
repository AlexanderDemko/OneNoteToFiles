using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using SongHelper.Services;
using System.Xml.Linq;
using OneNoteToFiles.Helpers;
using OneNoteToFiles.Consts;
using OneNoteToFiles.Services;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;

namespace OneNoteToFiles
{
    class Program
    {
        private static Application _oneNoteApp;        
        
        [STAThread]
        unsafe static void Main(string[] args)
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();
            Console.WriteLine("Start");

            try
            {
                _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

//                var s = @"<one:OE creationTime='2019-01-30T16:11:26.000Z' lastModifiedTime='2019-05-13T03:51:39.000Z' objectID='{E892A88C-B8DF-4EA0-85C7-60A9724D2E46}{17}{B0}' alignment='left' quickStyleIndex='1' style='font-family:Calibri;font-size:11.0pt'>
//                    <one:List>
//                      <one:Number numberSequence='58' numberFormat='##.' fontColor='#FF0000' fontSize='11.0' font='Calibri' italic='true' language='1049' text='а.' />
//                    </one:List>
//                    <one:T><![CDATA[<span
//style='font-style:italic;color:red' lang=ru>Аналогичная реакция была и у Петра, когда он осознал, Кто перед ним находится (</span><a
//href='isbtBibleVerse:rst/42%205:8;Луки%205:8'><span style='font-style:italic;
//background:#92D050' lang=ru>Лк 5:8</span></a><span style='font-style:italic;
//color:red' lang=en-US>). </span>]]></one:T>
//                  </one:OE>
//                  <one:OE creationTime='2019-01-30T16:07:54.000Z' lastModifiedTime='2019-05-13T03:51:39.000Z' objectID='{E892A88C-B8DF-4EA0-85C7-60A9724D2E46}{13}{B0}' alignment='left' quickStyleIndex='1' style='font-family:Calibri;font-size:11.0pt'>
//                    <one:List>
//                      <one:Number numberSequence='58' numberFormat='##.' fontColor='#FF0000' fontSize='11.0' font='Calibri' italic='true' language='1049' text='б.' />
//                    </one:List>
//                    <one:T><![CDATA[<span
//style='font-style:italic;color:red'>В конце своей жизни, находясь в ссылке на острове Патмос, Апостол Иоанн (наиболее приближённый к Иисусу Христу ученик) увидел Иисуса Христа. Иоанн не принялся Его обнимать, а упал как мёртвый (</span><a
//href='isbtBibleVerse:rst/66%201:17;Откровение%201:17&amp;qa=1'><span
//style='font-style:italic;background:#92D050'>Отк 1:17</span></a><span
//style='font-style:italic;color:red'> - дополнительная ссылка).</span>]]></one:T>
//                  </one:OE>";                

//                var result = SanitizeText(s);
//                Console.WriteLine(result);
                                
                SaveFiles();
            }
            catch (Exception ex)
            {
                Logger.LogError(ex.ToString());
            }

            sw.Stop();

            Console.WriteLine("Finish. Elapsed time: {0}", sw.Elapsed);
            Console.ReadKey();
        }

        private static void SaveFiles()
        {
            var fileExt = ".onenote";

            var pagesIds = GetPagesIds();
            var files = Directory.GetFiles(SettingsManager.TargetFolderPath, "*" + fileExt);
            foreach (var file in files)
                File.Delete(file);

            var index = 0;
            foreach (var pageId in pagesIds)
            {
                var pageContent = OneNoteUtils.GetPageContent(ref _oneNoteApp, pageId, out XmlNamespaceManager xnm);
                var pageName = (string)pageContent.Root.Attribute("name");
                pageName = $"{++index:D2}. {pageName}";

                Console.WriteLine($"Saving page: {pageName}");

                var filePath = Path.Combine(SettingsManager.TargetFolderPath, RemoveIllegalChars(pageName) + fileExt);
                var fileContent = pageContent.ToString();                

                if (SettingsManager.OnlyText)
                    fileContent = SanitizeText(fileContent);                
                else
                    fileContent = RemoveRedundantInfo(fileContent);

                File.WriteAllText(filePath, fileContent);
            }
        }

        private static string RemoveRedundantInfo(string s)
        {
            var attributesToDelete = new[]
            {
                "author",
                "authorInitials",
                "creationTime",
                "lastModifiedTime",
                "lastModifiedBy",
                "lastModifiedByInitials",
                "objectID",
                "quickStyleIndex",
                "lang",
                "objectID",
                "ID",
                "callbackID"
            };

            foreach (var attr in attributesToDelete)            
                s = Regex.Replace(s, $" {attr}=\"([^\"]*)\"", string.Empty);

            return s;
        }
        
        private static string SanitizeText(string s)
        {
            var sb = new StringBuilder();

            var startPattern = @"<![CDATA[";
            var startIndex = s.IndexOf(startPattern);
            while (startIndex > -1)
            {
                var endIndex = s.IndexOf("]]", startIndex);

                sb.AppendLine(s.Substring(startIndex + startPattern.Length, endIndex - startIndex - startPattern.Length));

                startIndex = s.IndexOf(startPattern, endIndex);
            }

            var result = sb.ToString();
            
            result = Regex.Replace(result, "<span[^>]*>", "<span>");
            result = Regex.Replace(result, "<a[^>]*>", "<a>");

            return result;
        }

        private static string RemoveIllegalChars(string path)
        {
            string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

            foreach (char c in invalid)
            {
                path = path.Replace(c.ToString(), "");
            }

            return path;
        }

        private static List<string> GetPagesIds()
        {
            var sectionPathParts = SettingsManager.SourceSectionPath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

            var notebookId = OneNoteUtils.GetNotebookIdByName(ref _oneNoteApp, sectionPathParts[0], false);
            string sectionGroupId = notebookId;            
            for (var i = 1; i < sectionPathParts.Length - 1; i++)            
                sectionGroupId = (string)OneNoteUtils.GetHierarchyElementByName(ref _oneNoteApp, "SectionGroup", sectionPathParts[i], sectionGroupId).Attribute("ID");                            

            var sectionId = (string)OneNoteUtils.GetHierarchyElementByName(ref _oneNoteApp, "Section", sectionPathParts.Last(), sectionGroupId).Attribute("ID");
            var sectionEl = ApplicationCache.Instance.GetHierarchy(ref _oneNoteApp, sectionId, HierarchyScope.hsPages);

            var pagesIds = new List<string>();
            foreach (var pageEl in sectionEl.Content.Root.Elements())
            {                
                pagesIds.Add((string)pageEl.Attribute("ID"));
            }

            return pagesIds;
        }
    }
}
