using System;
using System.Collections.Generic;
using System.Linq;
using MicroMWordLib.WordText;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordImage;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordOperations;

namespace MicroMWordLib.WordParagraph
{

    public enum Alignment
    {
        Left = 0,
        Center = 1,
        Right = 2,
        Both = 3
    }

    public enum ParagraphContent
    {
        Text = 1,
        Image = 2,
        Heading = 3,
        TextAndImage = 4
    }

    public enum ParagraphStyle
    {
        SimpleContainer = 0,
        ListItemElement = 1,
        WordHeading1 = 2,
        WordHeading2 = 3,
        WordHeading3 = 4,
        WordHeading4 = 5,
        WordHeading5 = 6,
        WordHeading6 = 7,
        WordHeading7 = 8,
        WordHeading8 = 9,
        WordHeading9 = 10
    }

    //public enum ParagraphStyle
    //{
    //    SimpleContainer = 0,
    //    WordHeading1 = 1,
    //    WordHeading2 = 2,
    //    WordHeading3 = 3,
    //    WordHeading4 = 4,
    //    WordHeading5 = 5,
    //    WordHeading6 = 6,
    //    WordHeading7 = 7,
    //    WordHeading8 = 8,
    //    WordHeading9 = 9
    //}

    public class WParagraph : IWBaseElement
    {
        private WCSelection prp_ContentSelection = null;
        private ParagraphContent prp_ParagraphContent = ParagraphContent.Text;
        private ParagraphStyle prp_ParagraphStyle = ParagraphStyle.SimpleContainer;
        private Alignment prp_Alingment = Alignment.Left;
        private int prp_ListItemLevel = -1;
        private int prp_ListID = -1;

        private List<IWParagraph> prp_Elements = null;

        public List<IWParagraph> Elements { get => prp_Elements; set => prp_Elements = value; }
        public Alignment Alingment { get => prp_Alingment; set => prp_Alingment = value; }
        public ParagraphContent ParagraphContent { get => prp_ParagraphContent; set => prp_ParagraphContent = value; }
        public ParagraphStyle ParagraphStyle { get => prp_ParagraphStyle; set => prp_ParagraphStyle = value; }
        public WCSelection ContentSelection { get => prp_ContentSelection; set => prp_ContentSelection = value; }

        public int ListID { get => prp_ListID; set => prp_ListID = value; }
        public int ListItemLevel { get => prp_ListItemLevel; set => prp_ListItemLevel = value; }

        public WParagraph()
        {
            prp_Elements = new List<IWParagraph>();
        }

        public string GetText()
        {
            List<WTextPart> AllTParts = new List<WTextPart>();
            foreach (IWParagraph iwel in this.prp_Elements)
            {
                if (iwel.GetType() == typeof(WTextPart))
                {
                    AllTParts.Add(iwel as WTextPart);
                }
            }
            ;
            return string.Join(string.Empty, AllTParts.Select(tp => tp.Text));
        }

        public string GetTextOfElement(int PElementIndex)
        {
            if (PElementIndex >= this.prp_Elements.Count)
            {
                return null;
            }

            if (this.prp_Elements[PElementIndex].GetType() != typeof(WTextPart))
            {
                return string.Empty;
            }

            WTextPart wtp = prp_Elements[PElementIndex] as WTextPart;

            if (wtp != null)
            {
                if (wtp.Text != null)
                {
                    return wtp.Text;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        public static WCSelection[] GetAllContentSelectionsForRange(Application MWordApp, Document MWordDocument, int RangeStart, int RangeEnd)
        {
            Document DraftDoc = MWordApp.Documents.Add(Visible: false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WCSelection> ParagWCS = new List<WCSelection>();

            Paragraphs RParag = DraftDoc.Range(RangeStart, RangeEnd).Paragraphs;

            for (int prg = 1; prg <= RParag.Count; prg++)
            {
                try
                {
                    string _tmpXML = RParag[prg].Range.XML;

                    if (_tmpXML == null || _tmpXML == string.Empty) continue;
                }
                catch
                {
                    continue;
                }

                WCSelection wcs = new WCSelection();

                wcs.ContentID = "RParagraph_" + prg;
                wcs.ContentSelectionStart = RParag[prg].Range.Start;
                wcs.ContentSelectionEnd = RParag[prg].Range.End;

                ParagWCS.Add(wcs);
            }

            for (int pindx = 0; pindx < ParagWCS.Count; pindx++)
            {
                ParagWCS[pindx].ContentID = "RParagraph_" + (pindx + 1);
            }

            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

            return ParagWCS.ToArray();
        }

        public static WParagraph RecoverInnerContentSelection(WParagraph in_WParagraph, int in_WCSelectionIndex)
        {

            bool updateWCSEnd = false;
            if (in_WParagraph.ContentSelection == null)
            {
                in_WParagraph.ContentSelection = new WCSelection() { ContentID = "WParagraph_" + Guid.NewGuid(), ContentSelectionStart = in_WCSelectionIndex };
                updateWCSEnd = true;
            }

            {
                foreach (IWParagraph iwparage in in_WParagraph.Elements)
                {
                    int pcntlength = 0;
                    if (iwparage.GetType() == typeof(WImage))
                    {
                        if ((iwparage as WImage).ContentSelection == null)
                        {
                            pcntlength++;
                        }
                        else
                        {
                            pcntlength += (iwparage as WImage).ContentSelection.ContentSelectionEnd - (iwparage as WImage).ContentSelection.ContentSelectionStart;
                        }
                    }
                    else if (iwparage.GetType() == typeof(WTextPart))
                    {
                        if ((iwparage as WTextPart).Text == null)
                        {
                            pcntlength++;
                        }
                        else
                        {
                            pcntlength = (iwparage as WTextPart).Text.Length;
                        }
                    }

                    in_WCSelectionIndex += pcntlength;
                }
            }

            if (updateWCSEnd == true)
            {
                in_WParagraph.ContentSelection.ContentSelectionEnd = in_WCSelectionIndex;
            }

            return in_WParagraph;
        }

        //public static WParagraph RecoverImages(WParagraph in_WParagraph, WImage[] in_WImages)
        //{
        //    for (int _imgIndex = 0; _imgIndex < in_WImages.Length; _imgIndex++)
        //    {
        //        if (in_WImages[_imgIndex].ContentSelection.ContentSelectionStart >= in_WParagraph.ContentSelection.ContentSelectionStart && in_WImages[_imgIndex].ContentSelection.ContentSelectionStart < in_WParagraph.ContentSelection.ContentSelectionEnd)
        //        {
        //            //int comSLengthPrgCmp = in_WImages[_imgIndex].ContentSelection.ContentSelectionStart;
        //            int comSLengthPrgCmp = in_WParagraph.ContentSelection.ContentSelectionStart;
        //            int _elemIndex = 0;
        //            while (in_WImages[_imgIndex].ContentSelection.ContentSelectionEnd > comSLengthPrgCmp)
        //            {
        //                if (_elemIndex < in_WParagraph.Elements.Count)
        //                {
        //                    string wptext = in_WParagraph.GetTextOfElement(_elemIndex);

        //                    wptext = (wptext == null) ? "1" : wptext;

        //                    comSLengthPrgCmp += wptext.Length;
        //                }
        //                else
        //                {
        //                    _elemIndex++;
        //                    break;
        //                }

        //                _elemIndex++;
        //            }

        //            if (comSLengthPrgCmp == in_WImages[_imgIndex].ContentSelection.ContentSelectionStart)
        //            {
        //                if (_elemIndex < in_WParagraph.Elements.Count)
        //                {
        //                    in_WParagraph.Elements.Insert(_elemIndex, (in_WImages[_imgIndex]));
        //                }
        //                else
        //                {
        //                    in_WParagraph.Elements.Add(in_WImages[_imgIndex]);
        //                }

        //                int ElemToDelete = 0;
        //                _elemIndex++;
        //                for (int rel = _elemIndex; rel < in_WParagraph.Elements.Count; rel++)
        //                {
        //                    string wptext = in_WParagraph.GetTextOfElement(_elemIndex);

        //                    wptext = (wptext == null) ? "1" : wptext;

        //                    comSLengthPrgCmp += wptext.Length;

        //                    if (comSLengthPrgCmp <= in_WImages[_imgIndex].ContentSelection.ContentSelectionEnd)
        //                    {
        //                        ElemToDelete++;
        //                    }
        //                    else
        //                    {
        //                        break;
        //                    }
        //                }

        //                if (ElemToDelete > 0)
        //                {
        //                    in_WParagraph.Elements.RemoveRange(_elemIndex, ElemToDelete);
        //                }
        //            }
        //            else
        //            {
        //                in_WParagraph.Elements.Insert(_elemIndex - 1, (in_WImages[_imgIndex]));
        //                _elemIndex--;
        //                int ElemToDelete = 0;
        //                for (int rel = _elemIndex; rel < in_WParagraph.Elements.Count; rel++)
        //                {
        //                    string wptext = in_WParagraph.GetTextOfElement(_elemIndex);

        //                    wptext = (wptext == null) ? "1" : wptext;

        //                    comSLengthPrgCmp += wptext.Length;

        //                    if (comSLengthPrgCmp <= in_WImages[_imgIndex].ContentSelection.ContentSelectionEnd)
        //                    {
        //                        ElemToDelete++;
        //                    }
        //                    else
        //                    {
        //                        break;
        //                    }
        //                }

        //                if (ElemToDelete > 0)
        //                {
        //                    in_WParagraph.Elements.RemoveRange(_elemIndex + 1, ElemToDelete);
        //                }
        //            }
        //        }
        //    }
        //    return in_WParagraph;
        //}

        public static WParagraph RecoverImages(WParagraph in_WParagraph, WImage[] in_WImages)
        {
            if (in_WParagraph.ContentSelection == null || in_WImages.Length < 1)
            {
                return in_WParagraph;
            }

            if (in_WParagraph.Elements.Count < 1)
            {
                return in_WParagraph;
            }

            int PAddedImageC = 0;

            for (int _imgIndex = 0; _imgIndex < in_WImages.Length; _imgIndex++)
            {
                if(in_WImages[_imgIndex].ContentSelection == null)
                {
                    continue;
                }

                if (in_WImages[_imgIndex].ContentSelection.ContentSelectionStart >= in_WParagraph.ContentSelection.ContentSelectionStart && in_WImages[_imgIndex].ContentSelection.ContentSelectionStart < in_WParagraph.ContentSelection.ContentSelectionEnd)
                {
                    for (int pe = PAddedImageC; pe < in_WParagraph.Elements.Count; pe++)
                    {
                        if (in_WParagraph.Elements[pe].GetType() == typeof(WImage))
                        {
                            in_WImages[_imgIndex].Width = (in_WParagraph.Elements[pe] as WImage).Width;
                            in_WImages[_imgIndex].Height = (in_WParagraph.Elements[pe] as WImage).Height;

                            in_WParagraph.Elements[pe] = in_WImages[_imgIndex];
                            PAddedImageC = pe + 1;

                            if(PAddedImageC >= in_WParagraph.Elements.Count)
                            {
                                return in_WParagraph;
                            }

                            break;
                        }
                    }
                }
            }
            return in_WParagraph;
        }

        public int RecoverImages(WImage[] in_WImages)
        {
            if (ContentSelection == null || in_WImages.Length < 1)
            {
                return 0;
            }

            if (in_WImages == null)
            {
                return 0;
            }

            if (Elements.Count < 1)
            {
                return 0;
            }

            int _AICount = 0;
            int PAddedImageC = 0;

            for (int _imgIndex = 0; _imgIndex < in_WImages.Length; _imgIndex++)
            {
                if (in_WImages[_imgIndex].ContentSelection.ContentSelectionStart >= ContentSelection.ContentSelectionStart && in_WImages[_imgIndex].ContentSelection.ContentSelectionStart <= ContentSelection.ContentSelectionEnd)
                {
                    for (int pe = PAddedImageC; pe < Elements.Count; pe++)
                    {
                        if (Elements[pe].GetType() == typeof(WImage))
                        {
                            in_WImages[_imgIndex].Width = (Elements[pe] as WImage).Width;
                            in_WImages[_imgIndex].Height = (Elements[pe] as WImage).Height;

                            Elements[pe] = in_WImages[_imgIndex];

                            _AICount++;
                            PAddedImageC = pe + 1;

                            if (PAddedImageC >= Elements.Count)
                            {
                                return _AICount;
                            }

                            break;
                        }
                    }
                }
            }
            return _AICount;
        }

        public static WParagraph[] RecoverImages(WParagraph[] in_WParagraphs, WImage[] in_WImages)
        {
            //int _LastSelection = -50;
            for (int _prgIndx = 0; _prgIndx < in_WParagraphs.Length; _prgIndx++)
            {
                if(in_WParagraphs[_prgIndx].ContentSelection == null)
                {
                    continue;
                }

                //if (in_WParagraphs[_prgIndx].ContentSelection.ContentSelectionStart - _LastSelection > 20)
                //{
                    in_WParagraphs[_prgIndx] = RecoverImages(in_WParagraphs[_prgIndx], in_WImages);
                //    _LastSelection = in_WParagraphs[_prgIndx].ContentSelection.ContentSelectionEnd;
                //}
            }
            return in_WParagraphs;
        }

        //public static WParagraph[] RecoverImages(WParagraph[] in_WParagraphs, WImage[] in_WImages)
        //{
        //    int _prgIndex = 0;
        //    int _prgLastIndex = 0;
        //    for (int _imgIndex = 0; _imgIndex < in_WImages.Length; _imgIndex++)
        //    {
        //        _prgIndex = _prgLastIndex;

        //        while (_prgIndex < in_WParagraphs.Length)
        //        {
        //            if (in_WParagraphs[_prgIndex] != null)
        //            {
        //                if (in_WImages[_imgIndex].ContentSelection.ContentSelectionStart >= in_WParagraphs[_prgIndex].ContentSelection.ContentSelectionStart && in_WImages[_imgIndex].ContentSelection.ContentSelectionStart < in_WParagraphs[_prgIndex].ContentSelection.ContentSelectionEnd)
        //                {
        //                    int comSLengthPrgCmp = in_WParagraphs[_prgIndex].ContentSelection.ContentSelectionStart;
        //                    int _elemIndex = 0;
        //                    while (in_WImages[_imgIndex].ContentSelection.ContentSelectionStart > comSLengthPrgCmp)
        //                    {
        //                        if (_elemIndex < in_WParagraphs[_prgIndex].Elements.Count)
        //                        {
        //                            string wptext = in_WParagraphs[_prgIndex].GetTextOfElement(_elemIndex);

        //                            wptext = (wptext == null) ? "1" : wptext;

        //                            comSLengthPrgCmp += wptext.Length;
        //                        }
        //                        else
        //                        {
        //                            _elemIndex++;
        //                            break;
        //                        }

        //                        _elemIndex++;
        //                    }

        //                    if (comSLengthPrgCmp == in_WImages[_imgIndex].ContentSelection.ContentSelectionStart)
        //                    {
        //                        if (_elemIndex < in_WParagraphs[_prgIndex].Elements.Count)
        //                        {
        //                            in_WParagraphs[_prgIndex].Elements.Insert(_elemIndex, (in_WImages[_imgIndex]));
        //                        }
        //                        else
        //                        {
        //                            in_WParagraphs[_prgIndex].Elements.Add(in_WImages[_imgIndex]);
        //                        }

        //                        int ElemToDelete = 0;
        //                        _elemIndex++;
        //                        for (int rel = _elemIndex; rel < in_WParagraphs[_prgIndex].Elements.Count; rel++)
        //                        {
        //                            string wptext = in_WParagraphs[_prgIndex].GetTextOfElement(_elemIndex);

        //                            wptext = (wptext == null) ? "1" : wptext;

        //                            comSLengthPrgCmp += wptext.Length;

        //                            if (comSLengthPrgCmp <= in_WImages[_imgIndex].ContentSelection.ContentSelectionEnd)
        //                            {
        //                                ElemToDelete++;
        //                            }
        //                            else
        //                            {
        //                                break;
        //                            }
        //                        }

        //                        if (ElemToDelete > 0)
        //                        {
        //                            in_WParagraphs[_prgIndex].Elements.RemoveRange(_elemIndex, ElemToDelete);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        in_WParagraphs[_prgIndex].Elements.Insert(_elemIndex - 1, (in_WImages[_imgIndex]));
        //                        _elemIndex--;
        //                        int ElemToDelete = 0;
        //                        for (int rel = _elemIndex; rel < in_WParagraphs[_prgIndex].Elements.Count; rel++)
        //                        {
        //                            string wptext = in_WParagraphs[_prgIndex].GetTextOfElement(_elemIndex);

        //                            wptext = (wptext == null) ? "1" : wptext;

        //                            comSLengthPrgCmp += wptext.Length;

        //                            if (comSLengthPrgCmp <= in_WImages[_imgIndex].ContentSelection.ContentSelectionEnd)
        //                            {
        //                                ElemToDelete++;
        //                            }
        //                            else
        //                            {
        //                                break;
        //                            }
        //                        }

        //                        if (ElemToDelete > 0)
        //                        {
        //                            in_WParagraphs[_prgIndex].Elements.RemoveRange(_elemIndex + 1, ElemToDelete);
        //                        }
        //                    }

        //                    _prgLastIndex = _prgIndex;
        //                }
        //            }
        //            _prgIndex++;
        //        }
        //    }

        //    return in_WParagraphs;
        //}
    }
}
