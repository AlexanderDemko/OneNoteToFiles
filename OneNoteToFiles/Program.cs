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
            
            foreach (var pageId in pagesIds)
            {
                var pageContent = OneNoteUtils.GetPageContent(ref _oneNoteApp, pageId, out XmlNamespaceManager xnm);
                var pageName = (string)pageContent.Root.Attribute("name");

                Console.WriteLine($"Saving page: {pageName}");

                File.WriteAllText(Path.Combine(SettingsManager.TargetFolderPath, RemoveIllegalChars(pageName) + fileExt), pageContent.ToString());                
            }
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
