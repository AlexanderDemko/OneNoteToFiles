using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml;
using OneNoteToFiles.Helpers;
using OneNoteToFiles.Common;
using OneNoteToFiles.Consts;

namespace OneNoteToFiles.Services
{
    /// <summary>
    /// Кэш OneNote
    /// </summary>
    public class ApplicationCache
    {

        #region Helper classes       

        public class OneNoteHierarchyContentId
        {
            public string ID { get; set; }
            public HierarchyScope ContentScope { get; set; }

            public override bool Equals(object obj)
            {
                return this.ID == ((OneNoteHierarchyContentId)obj).ID && this.ContentScope == ((OneNoteHierarchyContentId)obj).ContentScope;
            }

            public override int GetHashCode()
            {
                int result = this.ContentScope.GetHashCode();

                if (!string.IsNullOrEmpty(this.ID))
                    result = result ^ this.ID.GetHashCode();

                return result;
            }
        }

        public class HierarchyElement
        {
            public OneNoteHierarchyContentId Id { get; set; }
            public XDocument Content { get; set; }
            public XmlNamespaceManager Xnm { get; set; }
            public bool WasModified { get; set; }
        }

        #endregion

        private static readonly object _locker = new object();

        private static volatile ApplicationCache _instance = null;
        public static ApplicationCache Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_locker)
                    {
                        if (_instance == null)
                        {
                            _instance = new ApplicationCache();
                        }
                    }
                }

                return _instance;
            }
        }

        protected ApplicationCache()
        {

        }

        private Dictionary<OneNoteHierarchyContentId, HierarchyElement> _hierarchyContentCache = new Dictionary<OneNoteHierarchyContentId, HierarchyElement>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="hierarchyId"></param>
        /// <param name="scope"></param>
        /// <param name="refreshCache">Стоит ли загружать данные из OneNote (true) или из кэша (false)</param>
        /// <returns></returns>
        public HierarchyElement GetHierarchy(ref Application oneNoteApp, string hierarchyId, HierarchyScope scope, bool refreshCache = false)
        {
            OneNoteHierarchyContentId contentId = new OneNoteHierarchyContentId() { ID = hierarchyId, ContentScope = scope };

            HierarchyElement result;

            if (!_hierarchyContentCache.ContainsKey(contentId) || refreshCache)
            {
                lock (_locker)
                {
                    string xml = null;
                    try
                    {
                        OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                        {
                            oneNoteAppSafe.GetHierarchy(hierarchyId, scope, out xml, Constants.CurrentOneNoteSchema);
                        });
                    }
                    catch (Exception ex)
                    {
                        throw new HierarchyNotFoundException(string.Format("Не удаётся найти иерархию типа '{0}' для элемента '{1}': {2}", scope, hierarchyId, OneNoteUtils.ParseErrorAndMakeItMoreUserFriendly(ex.Message)));
                    }

                    XmlNamespaceManager xnm;
                    XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);

                    if (!_hierarchyContentCache.ContainsKey(contentId))
                        _hierarchyContentCache.Add(contentId, new HierarchyElement() { Id = contentId, Content = doc, Xnm = xnm });
                    else
                        _hierarchyContentCache[contentId].Content = doc;
                }
            }

            result = _hierarchyContentCache[contentId];

            return result;
        }
    }
}
