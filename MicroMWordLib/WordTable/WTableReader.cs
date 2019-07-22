using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordParagraph;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordList;
using MicroMWordLib.WordOperations;

namespace MicroMWordLib.WordTable
{
    public class WTableReader
    {
        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument)
        {
            Document DraftDoc = MWordApp.Documents.Add(Visible:false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WCSelection> TableSelections = new List<WCSelection>();

            for (int tbli = 1; tbli <= DraftDoc.Tables.Count; tbli++)
            {
                //// Log info
                //AddToLog("Working on table: " + tbli);
                Table CurTable = DraftDoc.Tables[tbli];

                //string TableXML = CurTable.Range.XML;

                //// Log info
                //AddToLog("Getting table content. Row count: " + CurTable.Rows.Count + "; Column count: " + CurTable.Columns.Count);

                //WTable wtbl = GetTableFromTableXMLData(TableXML);

                //for (int trow = 1; trow <= CurTable.Rows.Count; trow++)
                //{

                //    string[] TRow = new string[CurTable.Columns.Count];
                //    for(int tcol = 1; tcol <= CurTable.Columns.Count; tcol++)
                //    {

                //        try
                //        {

                //            TRow[tcol - 1] = CurTable.Cell(trow, tcol).Range.Text;
                //        }
                //        catch
                //        {
                //            TRow[tcol - 1] = string.Empty;
                //        }
                //    }
                //    TList_Tables.Add(TRow);
                //}

                //// Log info
                //AddToLog("All content collected...");

                CurTable.Select();

                WCSelection wcs = new WCSelection()
                {
                    ContentSelectionStart = MWordApp.Selection.Start,
                    ContentSelectionEnd = MWordApp.Selection.End
                };

                TableSelections.Add(wcs);

                //// Log info
                //AddToLog("Table was added as selection. Selection[Start, End] = [" + cs.RStart + ", " + cs.REnd + "]");
            }

            TableSelections = new List<WCSelection>(WCSelectionOperations.RemoveChilds(TableSelections.ToArray()));

            //{
            //    WCSelection[] ListSelections = WListReader.GetAllContentSelections(MWordApp, MWordDocument, Logger);

            //    TableSelections = new List<WCSelection>(WCSelectionOperations.RemoveChilds(TableSelections.ToArray(), ListSelections, Logger));

            //}

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

            {
                for (int ls = 0; ls < TableSelections.Count; ls++)
                {
                    TableSelections[ls].ContentID = "WTable_" + (ls + 1);
                }
            }

            return TableSelections.ToArray();
        }

        public static WTable[] GetAllTables(Application MWordApp, Document MWordDocument)
        {
            Document DraftDoc = MWordApp.Documents.Add();
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WTable> TList = new List<WTable>();
            WCSelection[] TSelection = GetAllContentSelections(MWordApp, MWordDocument);

            {
                WCSelection[] ListSelections = WListReader.GetAllContentSelections(MWordApp, MWordDocument);

                TSelection = WCSelectionOperations.RemoveChilds(TSelection, ListSelections);

            }


            for (int tbls = 0; tbls < TSelection.Length; tbls++)
            {
                //// Log info
                //AddToLog("Working on table: " + tbli);

                //// Log info
                //AddToLog("Getting table content. Row count: " + CurTable.Rows.Count + "; Column count: " + CurTable.Columns.Count);

                string TableXML = MWordDocument.Range(TSelection[tbls].ContentSelectionStart, TSelection[tbls].ContentSelectionEnd).XML;

                WTable wtbl = GetTableFromTableXMLData(TableXML);

                //for (int trow = 1; trow <= CurTable.Rows.Count; trow++)
                //{

                //    string[] TRow = new string[CurTable.Columns.Count];
                //    for(int tcol = 1; tcol <= CurTable.Columns.Count; tcol++)
                //    {

                //        try
                //        {

                //            TRow[tcol - 1] = CurTable.Cell(trow, tcol).Range.Text;
                //        }
                //        catch
                //        {
                //            TRow[tcol - 1] = string.Empty;
                //        }
                //    }
                //    TList_Tables.Add(TRow);
                //}

                //// Log info

                wtbl.ContentSelection = TSelection[tbls];

                TList.Add(wtbl);

                //// Log info
                //AddToLog("Table was added as selection. Selection[Start, End] = [" + cs.RStart + ", " + cs.REnd + "]");
            }

            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

            return TList.ToArray();
        }

        public static WTable GetTableFromTableXMLData(string XMLContent)
        {
            XmlDocument XMLDoc = new XmlDocument();
            XMLDoc.LoadXml(XMLContent);

            WTable WTable = new WTable()
            {
                TableName = "WTable_" + Guid.NewGuid().ToString()
            };

            //XmlReader MyXMLRdr = XmlReader.Create(XMLContent);

            // Getting table rows
            XmlNodeList AllTableRows = XMLDoc.GetElementsByTagName(WordXMLTags.WTN_Table)[0].ChildNodes;
            //XMLDoc.GetElementsByTagName(WordXMLTags.WordTagName_TableRow);

            List<bool> VMerge = new List<bool>();

            int RowCounter = 0;

            foreach (XmlNode trow in AllTableRows)
            {

                if(trow.Name != WordXMLTags.WordTagName_TableRow)
                {
                    continue;
                }

                // Getting table cells
                XmlNodeList TRowChilds = trow.ChildNodes;
                WTableRow WTRow = new WTableRow();

                int CellCounter = 0;

                foreach (XmlNode trchild in TRowChilds)
                {
                    //XmlNodeList CellChilds = cxmln.ChildNodes; 
                    if (trchild.Name == WordXMLTags.WordTagName_TableCell)
                    {
                        while (VMerge.Count <= CellCounter)
                        {
                            VMerge.Add(false);
                        }

                        int HMerge_CellCount = 1;

                        WTableCell WTCell = new WTableCell()
                        {
                            WordTableCellType = TableCellType.Normal
                        };
                        //WTCell.WordTableCellValueType = TableCellValueType.Empty;
                        //WTCell.WordTableCellValue = string.Empty;

                        if (RowCounter == 0)
                        {
                            VMerge.Add(false);
                        }

                        XmlDocument Cnts = new XmlDocument();
                        Cnts.LoadXml(trchild.OuterXml);

                        //XmlNodeList TableCellElements = Cnts.GetElementsByTagName(WordXMLTags.WordTagName_Paragraph);

                        //if (HMerge_CellCount > 1)
                        //{
                        //    WTCell.WordTableCellType = TableCellType.H_Merged;
                        //    HMerge_CellCount--;
                        //}
                        //else
                        //{
                        //    WTCell.WordTableCellType = TableCellType.Normal;
                        //}


                        WTCell.WordTableCellType = (VMerge[CellCounter] == true) ? TableCellType.V_Merged : TableCellType.Normal;
                        try
                        { 
                        if (trchild.ChildNodes != null)
                        {

                            List<IWBaseElement> WElemList = new List<IWBaseElement>();

                            foreach (XmlNode tcelem in trchild.ChildNodes)
                            {
                                if (tcelem.Name == WordXMLTags.WordTagName_Paragraph)
                                {
                                    WElemList.Add(WParagraphReader.GetParagraphFromParagraphXMLNode(tcelem));
                                }
                                else if (tcelem.Name == WordXMLTags.WTN_Table)
                                {
                                    WElemList.Add(GetTableFromTableXMLData(tcelem.OuterXml));
                                }
                            }


                            bool ListCrt = false;
                            //int LstIndex = 0;
                            int llistIndex = 0;
                            List<IWBaseElement> list_Elems = null;

                            for (int cel = 0; cel < WElemList.Count; cel++)
                            {
                                if (WElemList[cel].GetType() == typeof(WParagraph))
                                {
                                    WParagraph WPrg = WElemList[cel] as WParagraph;

                                    if (WPrg.ParagraphStyle != ParagraphStyle.ListItemElement)
                                    {
                                        WTCell.CellElements.Add(WPrg);
                                    }
                                    else
                                    {
                                        if (list_Elems == null)
                                        {
                                            list_Elems = new List<IWBaseElement>();
                                        }

                                        if (ListCrt == true)
                                        {
                                            if (WPrg.ListID == (list_Elems[0] as WParagraph).ListID)
                                            {
                                                if (WTCell.CellElements.Count - llistIndex > 1)
                                                {
                                                    IWBaseElement[] _tmpArr = new IWBaseElement[WTCell.CellElements.Count - llistIndex];
                                                    Array.Copy(WTCell.CellElements.ToArray(), llistIndex, _tmpArr, 0, _tmpArr.Length);

                                                    list_Elems.AddRange(_tmpArr);
                                                    WTCell.CellElements.RemoveRange(llistIndex, WTCell.CellElements.Count - llistIndex);
                                                }
                                                list_Elems.Add(WPrg);

                                            }
                                            else
                                            {
                                                WTCell.CellElements.Add(WListReader.GetListFromWordElements(list_Elems.ToArray()));

                                                list_Elems = new List<IWBaseElement>();
                                                llistIndex = WTCell.CellElements.Count - 1;
                                                ListCrt = true;
                                            }
                                        }
                                        else
                                        {
                                            list_Elems.Add(WPrg);
                                            llistIndex = WTCell.CellElements.Count - 1;
                                            ListCrt = true;
                                        }

                                        llistIndex = llistIndex < 0 ? 0 : llistIndex;
                                    }
                                }
                                else if (WElemList[cel].GetType() == typeof(WTable))
                                {
                                    WTCell.CellElements.Add(WElemList[cel]);
                                }

                                if (cel == WElemList.Count - 1)
                                {
                                    if (ListCrt == true)
                                    {
                                        WTCell.CellElements.Insert(llistIndex, WListReader.GetListFromWordElements(list_Elems.ToArray()));
                                    }
                                }
                            }
                            //foreach (XmlNode prg in TableCellElements)
                            //{ 
                            //    WTCell.AddParagraph(WParagraphReader.GetParagraphFromParagraphXMLNode(prg));
                            //}
                        }
                    }
                    catch(Exception Exp)
                    {

                    }
                        //if (WTCell.WordTableCellValue != null && WTCell.WordTableCellValue != string.Empty)
                        //{
                        //    WTCell.WordTableCellValue = WTCell.WordTableCellValue.Substring(0, WTCell.WordTableCellValue.LastIndexOf(WTable.SeparatorFor_Paragraph));
                        //}

                        XmlNode _GSpanNode = Cnts.GetElementsByTagName(WordXMLTags.WordTagName_GridSpan)[0];
                        XmlNode _VMergeNode = Cnts.GetElementsByTagName(WordXMLTags.WordTagName_VerticalMerge)[0];

                        if (_GSpanNode != null)
                        {
                            if (_GSpanNode.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                            {
                                HMerge_CellCount = Convert.ToInt32(_GSpanNode.Attributes[WordXMLTags.WordTagAttrib_Value].Value);

                                if (RowCounter == 0)
                                {
                                    VMerge.AddRange(Enumerable.Repeat<bool>(false, HMerge_CellCount - 1).ToArray());
                                }
                                else
                                {
                                    for (int bi = CellCounter + 1; bi < CellCounter + (HMerge_CellCount - 1); bi++)
                                    {
                                        VMerge[CellCounter] = false;
                                    }
                                }
                            }
                        }
                        else
                        {
                            HMerge_CellCount = 1;
                        }

                        if (_VMergeNode != null)
                        {
                            if (_VMergeNode.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                            {
                                if (_VMergeNode.Attributes[WordXMLTags.WordTagAttrib_Value].Value.ToLower() == WordXMLTags.WordTagAttribValue_Restart | _VMergeNode.Attributes[WordXMLTags.WordTagAttrib_Value].Value.ToLower() == WordXMLTags.WordTagAttribValue_Continue)
                                {
                                    VMerge[CellCounter] = true;
                                }
                            }
                        }
                        else
                        {
                            VMerge[CellCounter] = false;
                        }


                        WTRow.AddCell(WTCell);

                        CellCounter++;
                        CellCounter += HMerge_CellCount - 1;

                        Cnts.RemoveAll();

                        while (HMerge_CellCount > 1)
                        {
                            WTRow.AddCell(new WTableCell(TableCellType.H_Merged));
                            HMerge_CellCount--;
                        }
                    }
                    else if(trchild.Name == WordXMLTags.WordTagName_TableRow_Properties)
                    {
                        foreach(XmlNode prp in trchild.ChildNodes)
                        {
                            if(prp.Name == WordXMLTags.WordTagName_ConditionalFormatting)
                            {
                                if (prp.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                                {
                                    WTRow.ConditionalFormatting = prp.Attributes[WordXMLTags.WordTagAttrib_Value].Value;
                                }
                            }
                        }
                    }
                }

                WTable.AddRow(WTRow);
                RowCounter++;
            }

            XMLDoc.RemoveAll();
            VMerge.Clear();

            WTable.ArrangeTableCells();

            return WTable;
        }
    }

}
