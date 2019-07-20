using System.Configuration;

namespace OneNoteToFiles
{
    public static class SettingsManager
    {
        public static string SourceSectionPath { get; private set; }

        public static string TargetFolderPath { get; private set; }

        public static bool OnlyText { get; private set; }

        static SettingsManager()
        {
            SourceSectionPath = ConfigurationManager.AppSettings["SourceSectionPath"];
            TargetFolderPath = ConfigurationManager.AppSettings["TargetFolderPath"];
            OnlyText = bool.Parse(ConfigurationManager.AppSettings["OnlyText"]);
        }       
    }
}
