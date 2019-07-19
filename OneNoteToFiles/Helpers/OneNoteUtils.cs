using Microsoft.Office.Interop.OneNote;
using OneNoteToFiles.Consts;
using OneNoteToFiles.Services;
using SongHelper.Services;
using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OneNoteToFiles.Helpers
{
    public static class OneNoteUtils
    {
        public static string ParseErrorAndMakeItMoreUserFriendly(string exceptionMessage)
        {
            var originalHexValue = Regex.Match(exceptionMessage, @"0x800[A-F\d]+").Value;
            if (!string.IsNullOrEmpty(originalHexValue))
            {
                var hexValue = originalHexValue.Replace("0x", "FFFFFFFF");
                long decValue;
                if (long.TryParse(hexValue, System.Globalization.NumberStyles.HexNumber, CultureInfo.CurrentCulture.NumberFormat, out decValue))
                {
                    var errorCode = (Error)decValue;
                    var userFriendlyErrorMessage = GetUserFriendlyErrorMessage(errorCode);
                    exceptionMessage = exceptionMessage.Replace(originalHexValue, userFriendlyErrorMessage);
                }
            }

            return exceptionMessage;
        }

        public static Application CreateOneNoteAppSafe()
        {
            Application oneNoteApp = null;
            UseOneNoteAPIInternal(ref oneNoteApp, null, 0);
            return oneNoteApp;
        }


        private static void UseOneNoteAPIInternal(ref Application oneNoteApp, Action<IApplication> action, int attemptsCount)
        {
            try
            {
                if (oneNoteApp == null)
                    oneNoteApp = new Application();

                if (action != null)
                    action(oneNoteApp);
            }
            catch (COMException ex)
            {
                if (ex.Message.Contains("0x80010100")                                           // "System.Runtime.InteropServices.COMException (0x80010100): System call failed. (Exception from HRESULT: 0x80010100 (RPC_E_SYS_CALL_FAILED))";
                    || ex.Message.Contains("0x800706BA")
                    || ex.Message.Contains("0x800706BE")
                    || ex.Message.Contains("0x80010001")                                        // System.Runtime.InteropServices.COMException (0x80010001): Вызов был отклонен. (Исключение из HRESULT: 0x80010001 (RPC_E_CALL_REJECTED))
                    || ex.Message.Contains("0x80010108")                                        // RPC_E_DISCONNECTED
                    )
                {
                    Logger.LogMessageSilient("UseOneNoteAPI. Attempt {0}: {1}", attemptsCount, ex.Message);
                    if (attemptsCount <= 15)
                    {
                        attemptsCount++;
                        Thread.Sleep(1000 * attemptsCount);                        

                        ReleaseOneNoteApp(ref oneNoteApp);
                        UseOneNoteAPIInternal(ref oneNoteApp, action, attemptsCount);
                    }
                    else
                        throw;
                }
                else
                    throw;
            }
        }

        public static void ReleaseOneNoteApp(ref Application oneNoteApp)
        {
            if (oneNoteApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(oneNoteApp);
                }
                catch (Exception releaseEx)
                {
                    Logger.LogError(releaseEx);
                }

                oneNoteApp = null;
            }
        }

        public static XDocument GetHierarchyElement(ref Application oneNoteApp, string hierarchyId, HierarchyScope scope, out XmlNamespaceManager xnm)
        {
            string xml = null;
            UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.GetHierarchy(hierarchyId, scope, out xml, Constants.CurrentOneNoteSchema);
            });
            return GetXDocument(xml, out xnm);
        }

        public static XDocument GetXDocument(string xml, out XmlNamespaceManager xnm, bool setLineInfo = false)
        {
            var xd = !setLineInfo ? XDocument.Parse(xml) : XDocument.Parse(xml, LoadOptions.SetLineInfo);
            xnm = GetOneNoteXNM();
            return xd;
        }

        public static XmlNamespaceManager GetOneNoteXNM()
        {
            var xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", Constants.OneNoteXmlNs);

            return xnm;
        }

        private static string GetUserFriendlyErrorMessage(Error errorCode)
        {
            switch (errorCode)
            {
                case Error.hrPageReadOnly:
                    return "Страница доступна только для чтения.";
                case Error.hrInsertingInk:
                    return "Не удалось обновить страницу. Возможно на странице присутствуют нарисованные кистью элементы, которые на текущий момент не поддерживаются программой. Удалите такие элементы и повторите операцию.";
                default:
                    return errorCode.ToString();
            }
        }

        public static void UseOneNoteAPI(ref Application oneNoteApp, Action<IApplication> action)
        {
            UseOneNoteAPIInternal(ref oneNoteApp, action, 0);
        }

        public static void UpdatePageContentSafe(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm, bool repeatIfPageIsReadOnly = true)
        {
            UpdatePageContentSafeInternal(ref oneNoteApp, pageContent, xnm, repeatIfPageIsReadOnly ? (int?)0 : null);
        }

        private static void UpdatePageContentSafeInternal(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm, int? attemptsCount)
        {
            var inkNodes = pageContent.Root.XPathSelectElements("one:InkDrawing", xnm)
                            .Union(pageContent.Root.XPathSelectElements("//one:OE[.//one:InkDrawing]", xnm))
                            .Union(pageContent.Root.XPathSelectElements("one:Outline[.//one:InkWord]", xnm)).ToArray();

            foreach (var inkNode in inkNodes)
            {
                if (inkNode.XPathSelectElement(".//one:T", xnm) == null)
                    inkNode.Remove();
                else
                {
                    var inkWords = inkNode.XPathSelectElements(".//one:InkWord", xnm).Where(ink => ink.XPathSelectElement(".//one:CallbackID", xnm) == null).ToArray();
                    inkWords.Remove();
                }
            }

            var pageTitleEl = pageContent.Root.XPathSelectElement("one:Title", xnm);                // могли случайно удалить заголовок со страницы
            if (pageTitleEl != null && !pageTitleEl.HasElements && !pageTitleEl.HasAttributes)
                pageTitleEl.Remove();

            try
            {
                UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.UpdatePageContent(pageContent.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
                });
            }
            catch (COMException ex)
            {
                if (attemptsCount.GetValueOrDefault(int.MaxValue) < 30                                       // 15 секунд - но каждое обновление требует времени. поэтому на самом деле дольше
                    && (OneNoteUtils.IsError(ex, Error.hrPageReadOnly) || OneNoteUtils.IsError(ex, Error.hrSectionReadOnly)))
                {
                    Thread.Sleep(500);
                    UpdatePageContentSafeInternal(ref oneNoteApp, pageContent, xnm, attemptsCount + 1);
                }
                else
                    throw;
            }
        }

        public static bool IsError(Exception ex, Error error)
        {
            return ex.Message.IndexOf(error.ToString(), StringComparison.InvariantCultureIgnoreCase) > -1
                || ex.Message.IndexOf(GetHexError(error), StringComparison.InvariantCultureIgnoreCase) > -1;
        }

        private static string GetHexError(Error error)
        {
            return string.Format("0x{0}", Convert.ToString((int)error, 16));
        }

        public static XDocument GetPageContent(ref Application oneNoteApp, string pageId, out XmlNamespaceManager xnm)
        {
            return GetPageContent(ref oneNoteApp, pageId, PageInfo.piBasic, out xnm);
        }

        public static XDocument GetPageContent(ref Application oneNoteApp, string pageId, PageInfo pageInfo, out XmlNamespaceManager xnm)
        {
            string xml = null;

            UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.GetPageContent(pageId, out xml, pageInfo, Constants.CurrentOneNoteSchema);
            });

            return OneNoteUtils.GetXDocument(xml, out xnm);
        }

        public static void UpdateElementMetaData(XElement el, string key, string value, XmlNamespaceManager xnm)
        {
            var metaElement = el.XPathSelectElement(string.Format("one:Meta[@name=\"{0}\"]", key), xnm);
            if (metaElement != null)
            {
                metaElement.SetAttributeValue("content", value);
            }
            else
            {
                var nms = XNamespace.Get(Constants.OneNoteXmlNs);

                var meta = new XElement(nms + "Meta",
                                            new XAttribute("name", key),
                                            new XAttribute("content", value));

                var beforeMetaEl = el.XPathSelectElement("one:MediaPlaylist", xnm) ?? el.XPathSelectElement("one:PageSettings", xnm);

                if (beforeMetaEl != null)
                    beforeMetaEl.AddBeforeSelf(meta);
                else
                {
                    var afterMetaEl = el.XPathSelectElement("one:Tag", xnm);
                    if (afterMetaEl != null)
                        afterMetaEl.AddAfterSelf(meta);
                    else
                        el.AddFirst(meta);
                }
            }
        }

        public static XElement GetHierarchyElementByName(ref Application oneNoteApp, string elementTag, string elementName, string parentElementId)
        {
            XmlNamespaceManager xnm;
            var parentEl = GetHierarchyElement(ref oneNoteApp, parentElementId, HierarchyScope.hsChildren, out xnm);

            return parentEl.Root.XPathSelectElement(string.Format("one:{0}[@name=\"{1}\"]", elementTag, elementName), xnm);
        }

        public static string GetNotebookIdByName(ref Application oneNoteApp, string notebookName, bool refreshCache)
        {
            var hierarchy = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            var bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@name=\"{0}\"]", notebookName), hierarchy.Xnm);
            if (bibleNotebook == null)
                bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@nickname=\"{0}\"]", notebookName), hierarchy.Xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("ID");
            }

            return string.Empty;
        }
    }
}
