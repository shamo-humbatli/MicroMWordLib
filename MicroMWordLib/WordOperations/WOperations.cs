using System;
using System.Collections.Generic;
using WordApp = Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordImage;
using MicroMWordLib.WordList;
using MicroMWordLib.WordParagraph;
using MicroMWordLib.WordTable;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace MicroMWordLib.WordOperations
{
    public class WOperations
    {
        private WordApp.Application MyWordApplication = null;
        private WordApp.Document MyWordDocument = null;
        private string prp_ImageFileBaseName = "Images";
        private string MyWordFilePath = string.Empty;
        private string MyWorkingFolder = string.Empty;
        private Object ref_Missing = WParameters.Missing;

        private bool prp_JoinImages = false;

        public string WorkingFolder { get => MyWorkingFolder; set => MyWorkingFolder = value; }

        public string ImageFileBaseName { get => prp_ImageFileBaseName; set => prp_ImageFileBaseName = value; }
        public bool JoinImages { get => prp_JoinImages; set => prp_JoinImages = value; }

        public WOperations()
        {
  
        }

        public WOperations(string WordFilePath, string WFolder)
        {
            MyWordFilePath = WordFilePath;


            ImageFileBaseName = Path.GetFileNameWithoutExtension(WordFilePath);

            ImageFileBaseName = ImageFileBaseName.Trim();
            ImageFileBaseName = ImageFileBaseName.Replace(" ", string.Empty);

            ImageFileBaseName = (ImageFileBaseName.Length > 200) ? ImageFileBaseName.Substring(0, ImageFileBaseName.Length - 50) + "..." : ImageFileBaseName;

            MyWorkingFolder = WFolder;
        }



        public IWBaseElement[] GetWordElements()
        {

            MyWordApplication = new WordApp.Application();
            MyWordDocument = MyWordApplication.Documents.Open(MyWordFilePath, ReadOnly: false, Visible: false, OpenAndRepair: false, AddToRecentFiles: false, Revert: false, NoEncodingDialog: true, ConfirmConversions: false);

            try
            {
                //MyWordDocument = MyWordApplication.Documents.Open(MyWordFilePath, ref_Missing, false, false, ref_Missing, ref_Missing, ref_Missing, ref_Missing, ref_Missing, ref_Missing, ref_Missing, false, ref_Missing, ref_Missing, ref_Missing, ref_Missing);
  
 

                List<IWBaseElement> WElementsList = new List<IWBaseElement>();

                WParagraph[] WParags = WParagraphReader.GetAllParagraphs(MyWordApplication, MyWordDocument);

                WImage[] WImgs = null;
                if (JoinImages == true)
                {
                    WImgs = WImageExporter.ExportImages(MyWordApplication, MyWordDocument, WCSelectionOperations.JoinSelections(WImage.GetAllContentSelections(MyWordApplication, MyWordDocument), 10, "image_"), MyWorkingFolder, ImageFileBaseName);
                }
                else
                {
                    WImgs = WImageExporter.ExportImages(MyWordApplication, MyWordDocument, MyWorkingFolder, ImageFileBaseName);
                }

                WList[] WLists = WListReader.GetAllLists(MyWordApplication, MyWordDocument);

                WTable[] WTables = WTableReader.GetAllTables(MyWordApplication, MyWordDocument);

 
                for (int li = 0; li < WLists.Length; li++)
                {
                    int RCnt = WLists[li].RecoverInnerContentSelection(WParagraph.GetAllContentSelectionsForRange(MyWordApplication, MyWordDocument, WLists[li].ContentSelection.ContentSelectionStart, WLists[li].ContentSelection.ContentSelectionEnd));
                }

                for (int tbl = 0; tbl < WTables.Length; tbl++)
                {
                    int RCnt = WTables[tbl].RecoverInnerContentSelection(WParagraph.GetAllContentSelectionsForRange(MyWordApplication, MyWordDocument, WTables[tbl].ContentSelection.ContentSelectionStart, WTables[tbl].ContentSelection.ContentSelectionEnd));
                }

                WParags = WParagraph.RecoverImages(WParags, WImgs);
                WLists = WList.RecoverImages(WLists, WImgs);
                WTables = WTable.RecoverImages(WTables, WImgs);
             
                MyWordDocument.Close(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);
                MyWordApplication.Quit(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);

                int RRslt = System.Runtime.InteropServices.Marshal.ReleaseComObject(MyWordApplication);

                WElementsList.AddRange(WParags);
                WElementsList.AddRange(WLists);
                WElementsList.AddRange(WTables);

                return ArrangeInAscendingOrder(WElementsList.ToArray());
            }
            catch(Exception Exp)
            {
                MyWordDocument.Close(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);
                MyWordApplication.Quit(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);

                int RRslt = System.Runtime.InteropServices.Marshal.ReleaseComObject(MyWordApplication);
                return null;
            }
            finally
            {
  
            }
        }

        public static IWBaseElement[] ArrangeInAscendingOrder(IWBaseElement[] in_WElements)
        {
            List<IWBaseElement> _WElemList = new List<IWBaseElement>(in_WElements);
            _WElemList.Sort((a, b) => a.ContentSelection.ContentSelectionStart.CompareTo(b.ContentSelection.ContentSelectionStart));

            return _WElemList.ToArray();
        }
    }
}
