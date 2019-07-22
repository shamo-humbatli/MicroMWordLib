using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using System.Xml;
using MicroMWordLib.WordText;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordImage;
using MicroMWordLib.WordTable;
using MicroMWordLib.WordList;
using MicroMWordLib.WordOperations;
using MicroMWordLib.WordAdditionalElement;
using System.Threading;

namespace MicroMWordLib.WordParagraph
{
    public class WParagraphReader
    {
        private static char[] Numebers = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument)
        {
            try
            {
                Document DraftDoc = MWordApp.Documents.Add(Visible: false);
                MWordDocument.Select();
                MWordApp.Selection.Copy();
                DraftDoc.Range().Paste();
                DraftDoc.Activate();

                List<WCSelection> ParagWCS = new List<WCSelection>();


                for (int prg = 1; prg <= DraftDoc.Paragraphs.Count; prg++)
                {
                    //MyStringBuilder.Append(MyDraftDoc.Paragraphs[prg].Range.Text + Environment.NewLine);
                    WCSelection wcs = new WCSelection();

                    //wparag.ParagraphProperties = wparagp;
                    wcs.ContentID = "DParagraph_" + prg;
                    wcs.ContentSelectionStart = DraftDoc.Paragraphs[prg].Range.Start;
                    wcs.ContentSelectionEnd = DraftDoc.Paragraphs[prg].Range.End;

                    ParagWCS.Add(wcs);
                }


                WCSelection[] NParagWCS = WCSelectionOperations.RemoveChilds(ParagWCS.ToArray());

                for (int pindx = 0; pindx < NParagWCS.Length; pindx++)
                {
                    NParagWCS[pindx].ContentID = "WParagraph_" + (pindx + 1);
                }

                DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

                return NParagWCS;
            }
            catch(Exception Exp)
            {
                return null;
            }
        }

        public static WParagraph[] GetAllParagraphs(Application MWordApp, Document MWordDocument)
        {
            List<WParagraph> ListParagraphs = new List<WParagraph>();

            Document DraftDoc = MWordApp.Documents.Add(Visible: false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            //// Log info
            //AddToLog("Trying to get all paragraphs...");
            //if (MyDraftDoc.Paragraphs.Count < 1)
            //{
            //    AddToLog("No paragraphs found. Going to next step...");
            //}
            //else
            //{
            //    AddToLog("Paragraph count: " + MyDraftDoc.Paragraphs.Count);
            //}

            WCSelection[] AllPSelecs = GetAllContentSelections(MWordApp, MWordDocument);

            WCSelection[] TableWCS = WTableReader.GetAllContentSelections(MWordApp, MWordDocument);

            WCSelection[] ListWCS = WListReader.GetAllContentSelections(MWordApp, MWordDocument);

            WCSelection[] TOCWCS = WAElement.GetAllContentSelections(MWordApp, MWordDocument);


            //// Log
            //Logger.AddLog(new LogContent() { LogSubject = "Selection Array", LogMessage = "Getting all Image selections...", LogSeverity = LoggerParameters.LogSeverity.DEBUG });

            //WCSelection[] ImageSelections = WImage.GetAllContentSelections(MWordApp, MWordDocument, Logger);

            WCSelection[] AllPrgCSelections = WCSelectionOperations.RemoveCompairingParts(AllPSelecs, WCSelectionOperations.CreateNewArrangedSelectionArray(TableWCS, ListWCS));

            AllPrgCSelections = WCSelectionOperations.RemoveCompairingParts(AllPrgCSelections, TOCWCS);
           //AllPrgCSelections = WCSelectionOperations.RemoveAdditonalCompairingParts(AllPrgCSelections, TOCWCS, Logger);

           XmlDocument DraftXMLDoc = new XmlDocument();

            for (int psel = 0; psel < AllPrgCSelections.Length; psel++)
            {
                WParagraph wparag = new WParagraph();
                wparag.ContentSelection = AllPrgCSelections[psel];
                try
                {
                    //Range PRange = MWordDocument.Range(AllPrgCSelections[psel].ContentSelectionStart, AllPrgCSelections[psel].ContentSelectionEnd);
                    //string NLocal = PRange.get_Style().NameLocal;

                    string ParagXML = MWordDocument.Range(AllPrgCSelections[psel].ContentSelectionStart, AllPrgCSelections[psel].ContentSelectionEnd).XML;
                    DraftXMLDoc.LoadXml(ParagXML);
                    string tmpBody = DraftXMLDoc.GetElementsByTagName(WordXMLTags.WTN_Body)[0].OuterXml;
                    DraftXMLDoc.RemoveAll();
                    DraftXMLDoc.LoadXml(tmpBody);
                    XmlNode PrgNd = DraftXMLDoc.GetElementsByTagName(WordXMLTags.WordTagName_Paragraph)[0];
                    wparag = GetParagraphFromParagraphXMLNode(PrgNd);
                    wparag.ContentSelection = AllPrgCSelections[psel];
                    ListParagraphs.Add(wparag);
                    DraftXMLDoc.RemoveAll();
                }
                catch(Exception Exp)
                {
                    //ListParagraphs.Add(wparag);

                    //// Log
                    //Logger.AddLog(new LogContent() { LogSubject = "Getting Paragraph", LogMessage = "Paragraph [" + (psel + 1) + "/" + "] added. Content selection -> [" + AllPrgCSelections[psel].ContentSelectionStart + " : " + AllPrgCSelections[psel].ContentSelectionEnd + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });
                }
            }


            //for (int prg = 1; prg <= DraftDoc.Paragraphs.Count; prg++)
            //{
            //    //MyStringBuilder.Append(MyDraftDoc.Paragraphs[prg].Range.Text + Environment.NewLine);
            //    WCSelection wcs = new WCSelection();
            //    WParagraph wparag = new WParagraph();
            //    WParagraphProperties wparagp = new WParagraphProperties();

            //    //wparag.ParagraphProperties = wparagp;
            //    wcs.ContentID = "Paragraph_" + prg;
            //    wcs.ContentSelectionStart = DraftDoc.Paragraphs[prg].Range.Start;
            //    wcs.ContentSelectionEnd = DraftDoc.Paragraphs[prg].Range.End;

            //    XmlDocument DraftXMLDoc = new XmlDocument();

            //    try
            //    {
            //        {
            //            //if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading1].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading1;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading2].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading2;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading3].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading3;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading4].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading4;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading5].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading5;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading6].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading6;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading7].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading7;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading8].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading8;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleHeading9].NameLocal)
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading9;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal.ToString().ToLower.Contains("heading"))
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordHeading9;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Range.get_Style().NameLocal == DraftDoc.Styles[WdBuiltinStyle.wdStyleListParagraph].NameLocal)
            //            //{
            //            //    break;
            //            //}
            //            //else
            //            //{
            //            //    wparagp.ParagraphStyle = ParagraphStyle.WordSimpleParagraph;
            //            //}

            //            //if (DraftDoc.Paragraphs[prg].Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
            //            //{
            //            //    wparagp.Alingment = Alignment.Center;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Alignment == WdParagraphAlignment.wdAlignParagraphLeft)
            //            //{
            //            //    wparagp.Alingment = Alignment.Left;
            //            //}
            //            //else if (DraftDoc.Paragraphs[prg].Alignment == WdParagraphAlignment.wdAlignParagraphRight)
            //            //{
            //            //    wparagp.Alingment = Alignment.Right;
            //            //}
            //            //else
            //            //{
            //            //    wparagp.Alingment = Alignment.Both;
            //            //}
            //        }

            //        string ParagXML = DraftDoc.Paragraphs[prg].Range.XML;
            //        DraftXMLDoc.LoadXml(ParagXML);

            //        XmlNode PrgNd = DraftXMLDoc.GetElementsByTagName(WordXMLTags.WordTagName_Paragraph)[0];
            //        wparag = GetParagraphFromParagraphXMLNode(PrgNd);

            //        wparag.ParagraphProperties.ContentSelection = wcs;

            //        ListParagraphs.Add(wparag);
            //        DraftXMLDoc.RemoveAll();
            //    }
            //    catch
            //    {
            //        wparagp.ContentSelection = wcs;
            //        wparag.ParagraphProperties = wparagp;
            //        ListParagraphs.Add(wparag);
            //        DraftXMLDoc.RemoveAll();
            //    }


            //    //// Log info
            //    //AddToLog("Paragraph [" + prg + "/" + MyDraftDoc.Paragraphs.Count + "] -> Selection[Start, End] = " + "[" + Selcs[0] + ", " + Selcs[1] + "]");
            //}

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
            return ListParagraphs.ToArray();
        }

        public static WParagraph GetParagraphFromParagraphXMLNode(XmlNode PragraphNode)
        {
            if (PragraphNode != null)
            {
                WParagraph WPrg = new WParagraph();
                WPrg.ParagraphStyle = ParagraphStyle.SimpleContainer;
                XmlDocument PrgXML = new XmlDocument();
                PrgXML.LoadXml(PragraphNode.OuterXml);

                XmlNodeList wtextruns = PrgXML.GetElementsByTagName(WordXMLTags.WordTagName_TextRun);

                XmlNode pprps = PrgXML.GetElementsByTagName(WordXMLTags.WordTagName_Paragraph_Properties)[0];
                bool isHeader = false;
            
                if (pprps != null)
                {
                    foreach (XmlNode pprp in pprps.ChildNodes)
                    {
                        if (pprp.Name == WordXMLTags.WordTagName_Paragraph_Properties_Alignment)
                        {
                            if (pprp.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                            {
                                string aling = pprp.Attributes[WordXMLTags.WordTagAttrib_Value].Value;

                                if (aling.ToLower() == Alignment.Left.ToString().ToLower())
                                {
                                    WPrg.Alingment = Alignment.Left;
                                }
                                else if (aling.ToLower() == Alignment.Right.ToString().ToLower())
                                {
                                    WPrg.Alingment = Alignment.Right;
                                }
                                else if (aling.ToLower() == Alignment.Center.ToString().ToLower())
                                {
                                    WPrg.Alingment = Alignment.Center;
                                }
                                else if (aling.ToLower() == Alignment.Both.ToString().ToLower())
                                {
                                    WPrg.Alingment = Alignment.Both;
                                }
                                else
                                {
                                    WPrg.Alingment = Alignment.Left;
                                }

                            }
                        }
                        else if (pprp.Name == WordXMLTags.WTN_Pgraph_Properties_Style)
                        {
                            if (pprp.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                            {
                                string pstyle = pprp.Attributes[WordXMLTags.WordTagAttrib_Value].Value;
                                WPrg.ParagraphStyle = ParagraphStyle.SimpleContainer;
                                if (pstyle.ToLower() == "heading1")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading1;
                                }
                                else if (pstyle.ToLower() == "heading2")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading2;
                                }
                                else if (pstyle.ToLower() == "heading3")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading3;
                                }
                                else if (pstyle.ToLower() == "heading4")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading4;
                                }
                                else if (pstyle.ToLower() == "heading5")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading5;
                                }
                                else if (pstyle.ToLower() == "heading6")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading6;
                                }
                                else if (pstyle.ToLower() == "heading7")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading7;
                                }
                                else if (pstyle.ToLower() == "heading8")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading8;
                                }
                                else if (pstyle.ToLower() == "heading9")
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading9;
                                }
                                else if(pstyle.ToLower().Contains("heading"))
                                {
                                    WPrg.ParagraphStyle = ParagraphStyle.WordHeading9;
                                }

                                if (pstyle.ToLower().Contains("heading"))
                                {
                                    isHeader = true;
                                }

                            }
                        }
                        else if (pprp.Name == WordXMLTags.WTN_Pgraph_LPrp && isHeader == false)
                        {
                            string ilvl = null;
                            string ilfo = null;

                            foreach (XmlNode lprp in pprp.ChildNodes)
                            {
                                if (lprp.Name == WordXMLTags.WTN_Pgraph_LPrp_LLvl1)
                                {
                                    if (lprp.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                                    {
                                        ilvl = lprp.Attributes[WordXMLTags.WordTagAttrib_Value].Value;
                                    }
                                }
                                else if (lprp.Name == WordXMLTags.WTN_Pgraph_LPrp_LLvl2)
                                {
                                    if (lprp.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                                    {
                                        ilfo = lprp.Attributes[WordXMLTags.WordTagAttrib_Value].Value;
                                    }
                                }
                            }

                            if (ilfo != null)
                            {
                                WPrg.ListID = Convert.ToInt32(ilfo);
                            }

                            if (ilvl != null)
                            {
                                WPrg.ListItemLevel = Convert.ToInt32(ilvl);
                            }

                            WPrg.ParagraphStyle = ParagraphStyle.ListItemElement;
                        }
                    }
                }

                foreach (XmlNode WTextRun in wtextruns)
                {
                    XmlDocument TRun = new XmlDocument();
                    TRun.LoadXml(WTextRun.OuterXml);

                    XmlNode runtext = TRun.GetElementsByTagName(WordXMLTags.WordTagName_Text)[0];
                    XmlNode runprps = TRun.GetElementsByTagName(WordXMLTags.WordTagName_TextRun_Properties)[0];

                    WTextPart tpart = new WTextPart();

                    if(TRun.GetElementsByTagName(WordXMLTags.WTN_Picture)[0] != null)
                    {
                        string[] AVlArr = GetAtribValue(TRun.GetElementsByTagName(WordXMLTags.WTN_Picture)[0], WordXMLTags.WTA_Style).Split(';');

                        int rlWidth = -1, rlHeight = -1;

                        if (AVlArr != null)
                        {
                            if(AVlArr.Length > 0)
                            {
                                string rsWidth = Array.Find(AVlArr, ln => ln.ToLower().StartsWith("width")).ToString();
                                string rsHeight = Array.Find(AVlArr, ln => ln.ToLower().StartsWith("height")).ToString();

                                if (string.IsNullOrEmpty(rsWidth) != true)
                                {
                                    rsWidth = rsWidth.Substring("width".Length);

                                    string tmpsWidth = string.Empty;

                                    foreach(char se in rsWidth)
                                    {
                                        if(Array.IndexOf(Numebers, se) >= 0)
                                        {
                                            tmpsWidth += se;
                                        }
                                        else if(se == '.' || se == ',')
                                        {
                                            break;
                                        }
                                    }

                                    try
                                    {
                                        rlWidth = int.Parse(tmpsWidth);
                                    }
                                    catch
                                    {
                                        rlWidth = -1;
                                    }
                                }

                                if (string.IsNullOrEmpty(rsHeight) != true)
                                {
                                    rsWidth = rsWidth.Substring("height".Length);

                                    string tmpsHeight = string.Empty;

                                    foreach (char se in rsHeight)
                                    {
                                        if (Array.IndexOf(Numebers, se) >= 0)
                                        {
                                            tmpsHeight += se;
                                        }
                                        else if (se == '.' || se == ',')
                                        {
                                            break;
                                        }
                                    }

                                    try
                                    {
                                        rlHeight = int.Parse(tmpsHeight);
                                    }
                                    catch
                                    {
                                        rlHeight = -1;
                                    }
                                }
                            }
                        }


                        WPrg.Elements.Add(new WImage() { Width = rlWidth, Height = rlHeight });

                    }

                    tpart.Text = (runtext != null) ? runtext.InnerText : null;

                    if (runprps != null)
                    {
                        foreach (XmlNode rprp in runprps.ChildNodes)
                        {
                            if (rprp.Name == WordXMLTags.WordTagName_TextRun_Properties_Bold)
                            {
                                tpart.Bold = true;
                            }

                            if (rprp.Name == WordXMLTags.WordTagName_TextRun_Properties_Underline)
                            {
                                tpart.Underline = true;
                            }

                            if (rprp.Name == WordXMLTags.WordTagName_TextRun_Properties_Italic)
                            {
                                tpart.Italic = true;
                            }
                        }
                    }

                    WPrg.Elements.Add(tpart);
                }

                //if (WTCell.WordTableCellValue != null && WTCell.WordTableCellValue != string.Empty)
                //{
                //    WTCell.WordTableCellValue += WTable.SeparatorFor_Paragraph;
                //}

                PrgXML.RemoveAll();
                return WPrg;
            }
            return null;
        }

        private static string GetAtribValue(XmlNode in_XmlNode, string in_AtribName)
        {
            string out_Value = null;

            if (in_XmlNode.Attributes != null)
            {
                if (in_XmlNode.Attributes[in_AtribName] != null)
                {
                    out_Value = in_XmlNode.Attributes[in_AtribName].Value;
                }
            }

            if(out_Value == null)
            {
                if(in_XmlNode.ChildNodes != null)
                {
                    foreach(XmlNode cxmln in in_XmlNode.ChildNodes)
                    {
                        if(cxmln.NodeType != XmlNodeType.Element)
                        {
                            continue;
                        }
                        string AVl = GetAtribValue(cxmln, in_AtribName);
                        if(AVl != null)
                        {
                            return AVl;
                        }
                    }
                }
            }

            return out_Value;
        }
    }
}
