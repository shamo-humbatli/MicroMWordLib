using System;
using System.Collections.Generic;
using System.Xml;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordParagraph;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordTable;
using MicroMWordLib.WordOperations;
namespace MicroMWordLib.WordList
{
    public class WListReader
    {

        public WListReader()
        {

        }

        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument)
        {
            Document DraftDoc = MWordApp.Documents.Add(Visible: false);
            //Document DraftDoc = MWordApp.Documents.Add(null, null, null, false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WCSelection> ListSelection = new List<WCSelection>();

            for (int wlist = 1; wlist <= DraftDoc.Lists.Count; wlist++)
            {
                List dlist = DraftDoc.Lists[wlist];
                dlist.Range.Select();
                WCSelection wcs = new WCSelection()
                {
                    ContentSelectionStart = MWordApp.Selection.Start,
                    ContentSelectionEnd = MWordApp.Selection.End
                };
                ListSelection.Add(wcs);
            }


            ListSelection = new List<WCSelection>(WCSelectionOperations.RemoveChilds(ListSelection.ToArray()));

            //{
            //    WCSelection[] TableSelections = WTableReader.GetAllContentSelections(MWordApp, MWordDocument, Logger);

            //    ListSelection = new List<WCSelection>(WCSelectionOperations.RemoveChilds(ListSelection.ToArray(), TableSelections, Logger));

            //}

            {
                for(int ls = 0; ls < ListSelection.Count; ls++)
                {
                    ListSelection[ls].ContentID = "WList_" + (ls + 1);
                }
            }

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
            return ListSelection.ToArray();
        }

        public static WList[] GetAllLists(Application MWordApp, Document MWordDocument)
        {

            Document DraftDoc = MWordApp.Documents.Add(Visible: false);
            //Document DraftDoc = MWordApp.Documents.Add(null, null, null, false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WList> AllLists = new List<WList>();
            WCSelection[] ListSelection = GetAllContentSelections(MWordApp, MWordDocument);

            {
                WCSelection[] TableSelections = WTableReader.GetAllContentSelections(MWordApp, MWordDocument);

                ListSelection = WCSelectionOperations.RemoveChilds(ListSelection, TableSelections);

            }

            {
                int _csIndex = 0;
                int _lsCount = ListSelection.Length;

                while (_csIndex < _lsCount)
                {
                    string ListXML = null;
                    try
                    {
                        ListXML = MWordDocument.Range(ListSelection[_csIndex].ContentSelectionStart, ListSelection[_csIndex].ContentSelectionEnd).XML;

                        if(string.IsNullOrEmpty(ListXML) == true)
                        {
                            _csIndex++;
                            continue;
                        }
                    }
                    catch
                    {
                        _csIndex++;
                        continue;
                    }
                    XmlDocument DrftDocFullCntx = new XmlDocument();
                    DrftDocFullCntx.LoadXml(ListXML);

                    WList LContent = GetListFromListXMLData((DrftDocFullCntx.GetElementsByTagName(WordXMLTags.WTN_Body)[0]).ChildNodes[0]);
                    LContent.ContentSelection = ListSelection[_csIndex];
                    AllLists.Add(LContent);
                    _csIndex++;

                }

            }

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

            return AllLists.ToArray();
        }

        public static WList GetListFromListXMLData(XmlNode in_XMLNode)
        {
            //XmlNodeList Prgs = GetAllNodes(XMLContent);
            List<WListItemElement> PListItems = new List<WListItemElement>();

            foreach (XmlNode P in in_XMLNode.ChildNodes)
            {
                WListItemElement tmpWLIElem = GetListItemElementFromXMLData(P);
                if (tmpWLIElem.ListItemElement != null)
                {
                    PListItems.Add(tmpWLIElem);
                }
            }
            WList out_Rslt = new WList();
            PrepareListRecursively(out_Rslt, PListItems.ToArray());
            return out_Rslt;
        }

        public static WList GetListFromWordElements(IWBaseElement[] ListElements)
        {
            if (ListElements == null)
            {
                return null;
            }

            WListItemElement[] WLIElems = new WListItemElement[ListElements.Length];

            for (int pi = 0; pi < ListElements.Length; pi++)
            {
                WLIElems[pi] = new WListItemElement();
                if (ListElements[pi].GetType() == typeof(WParagraph))
                {
                    WLIElems[pi].ListItemElement = ListElements[pi] as WParagraph;
                    WLIElems[pi].ListID = (ListElements[pi] as WParagraph).ListID;
                    WLIElems[pi].ListItemLevel = (ListElements[pi] as WParagraph).ListItemLevel;
                }
                else if (ListElements[pi].GetType() == typeof(WTable))
                {
                    WLIElems[pi].ListItemElement = ListElements[pi] as WTable;
                    WLIElems[pi].ListID = -1;
                    WLIElems[pi].ListItemLevel = -1;
                }
            }

            WList out_Rslt = new WList();
            PrepareListRecursively(out_Rslt, WLIElems);
            return out_Rslt;

        }

        protected static int PrepareListRecursively(WList out_WList, WListItemElement[] ListItemElements)
        {

            if (ListItemElements == null)
            {
                return 0;
            }


            if (ListItemElements.Length < 1)
            {
                return 0;
            }

            //WList out_WList = new WList();
            int Level = ListItemElements[0].ListItemLevel;
            int LIIndex = 0;
            while (LIIndex < ListItemElements.Length)
            {
                WListItem wli = new WListItem();
               
                if (ListItemElements[LIIndex].ListItemLevel > Level)
                {
                    WListItemElement[] nplist = new WListItemElement[ListItemElements.Length - LIIndex];

                    Array.Copy(ListItemElements, LIIndex, nplist, 0, nplist.Length);
                    WList tmpCWList = new WList();
                    LIIndex += PrepareListRecursively(tmpCWList, nplist);

                    out_WList.ListItems[out_WList.ListItems.Count - 1].ChildList = tmpCWList;

                    continue;
                }
                else if (ListItemElements[LIIndex].ListItemLevel < Level)
                {
                    if(ListItemElements[LIIndex].ListItemLevel < 0)
                    {
                        out_WList.ListItems[out_WList.ListItems.Count - 1].ItemElements.Add(ListItemElements[LIIndex]);
                        LIIndex++;
                        continue;
                    }
                    else return LIIndex;
                }
                else
                {
                    wli.ListID = ListItemElements[LIIndex].ListID;
                    wli.ListItemLevel = ListItemElements[LIIndex].ListItemLevel;
                    wli.ItemElements.Add(ListItemElements[LIIndex]);
                }

                out_WList.ListItems.Add(wli);
                LIIndex++;
            }

            return LIIndex;
        }

        private static XmlNodeList GetAllNodes(string XMLContent)
        {
            try
            {
                XmlDocument AllPrgs = new XmlDocument();
                AllPrgs.LoadXml(XMLContent);
            
                //XmlNodeList prgs = AllPrgs.GetElementsByTagName(WordXMLTags.WordTagName_Paragraph);
                XmlNodeList prgs = AllPrgs.GetElementsByTagName(WordXMLTags.WTN_Body)[0].ChildNodes;
                return prgs;
            }
            catch
            {
                return null;
            }
        }

        private static WListItemElement GetListItemElementFromXMLData(XmlNode ListItemNode)
        {
            if (ListItemNode != null)
            {
                WListItemElement wlie = new WListItemElement();

                if (ListItemNode.Name == WordXMLTags.WordTagName_Paragraph)
                {
                    WParagraph WPrg = WParagraphReader.GetParagraphFromParagraphXMLNode(ListItemNode);
                    wlie.ListItemElement = WPrg;
                    wlie.ListID = WPrg.ListID;
                    wlie.ListItemLevel = WPrg.ListItemLevel;
                }
                else if (ListItemNode.Name == WordXMLTags.WTN_Table)
                {
                    WTable LTable = WTableReader.GetTableFromTableXMLData(ListItemNode.OuterXml);
                    wlie.ListItemElement = LTable;
                    wlie.ListItemLevel = -1;
                    wlie.ListID = -1;
                }

                return wlie;
            }
            return null;
        }
    }
}
