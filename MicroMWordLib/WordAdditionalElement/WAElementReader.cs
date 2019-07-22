using System;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using System.Xml;
using MicroMWordLib.WordText;
using MicroMWordLib.WordAdditionalElement.WordTOContents;
using MicroMWordLib.WordAdditionalElement.WordTOFigures;

namespace MicroMWordLib.WordAdditionalElement
{
    public class WAElementReader
    {
        public WAElement GetWAElement(Application MWordApp, Document MWordDocument, WAElement.WAElementType AdditionalElementType)
        {
            WAElement out_WAElement = null;

            Document DraftDoc = MWordApp.Documents.Add();
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            WCSelection[] WAElementSelection;
            if (AdditionalElementType == WAElement.WAElementType.TableOfContents)
            {
                WAElementSelection = WTOContents.GetAllContentSelections(MWordApp, MWordDocument);
            }
            else
            {
                WAElementSelection = WTOFigures.GetAllContentSelections(MWordApp, MWordDocument);
            }

            foreach (WCSelection WCSel in WAElementSelection)
            {
                out_WAElement = GetWAElementFromXML(MWordDocument.Range(WCSel.ContentSelectionStart, WCSel.ContentSelectionEnd).XML);
                out_WAElement.ContentSelection = WCSel;
            }

            return out_WAElement;
        }

        public static WAElement GetWAElementFromXML(string in_XMLContent)
        {
            try
            {
                WAElement out_WAElement = new WAElement();
                XmlDocument InitialDoc = new XmlDocument();
                InitialDoc.LoadXml(in_XMLContent);

                XmlNodeList AllParags = InitialDoc.GetElementsByTagName(WordXMLTags.WordTagName_Paragraph)[0].ChildNodes;

                WAElementLine tocLine = new WAElementLine();

                for (int prg = 0; prg < AllParags.Count; prg++)
                {
                    XmlDocument DrftXMLDoc = new XmlDocument();
                    DrftXMLDoc.LoadXml(AllParags[prg].OuterXml);

                    XmlNodeList AllPTRuns = DrftXMLDoc.GetElementsByTagName(WordXMLTags.WTN_Hyperlink)[0].ChildNodes;

                    WAElementLineField lineField = new WAElementLineField();

                    foreach (XmlNode WTextRun in AllPTRuns)
                    {
                        XmlDocument TRun = new XmlDocument();
                        TRun.LoadXml(WTextRun.OuterXml);

                        XmlNode runtext = TRun.GetElementsByTagName(WordXMLTags.WordTagName_Text)[0];
                        XmlNode runprps = TRun.GetElementsByTagName(WordXMLTags.WordTagName_TextRun_Properties)[0];

                        WTextPart tpart = new WTextPart();

                        if (runtext != null)
                        {
                            tpart.Text = runtext.InnerText;

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

                            lineField.Elements.Add(tpart);
                        }
                        else
                        {
                            lineField = new WAElementLineField();
                        }
                    }

                    tocLine.Fields.Add(lineField);
                }

                out_WAElement.Lines.Add(tocLine);
                return out_WAElement;
            }
            catch (Exception Exp)
            {
                return null;
            }
        }
    }
}
