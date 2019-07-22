using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordOperations;

namespace MicroMWordLib.WordImage
{
    public class WImage : IWParagraph
    {
        private WCSelection prp_ContentSelection = null;
        private string prp_WImagePath = null;
        private int prp_Width = -1;
        private int prp_Height = -1;


        public string ImagePath { get => prp_WImagePath; set => prp_WImagePath = value; }
        public WCSelection ContentSelection { get => prp_ContentSelection; set => prp_ContentSelection = value; }

        public int Height { get => prp_Height; set => prp_Height = value; }
        public int Width { get => prp_Width; set => prp_Width = value; }

        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument)
        {
            List<WCSelection> DraftIShapes = new List<WCSelection>();
            List<WCSelection> DraftShapes = new List<WCSelection>();

            Document DraftDoc = MWordApp.Documents.Add(Visible: true);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            for (int IShpI = 1; IShpI <= DraftDoc.InlineShapes.Count; IShpI++)
            {
                InlineShape IShape = DraftDoc.InlineShapes[IShpI];
                IShape.Select();

                WCSelection wcs = new WCSelection();
                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;
                DraftIShapes.Add(wcs);
            }

            for (int ShpI = 1; ShpI <= DraftDoc.Shapes.Count; ShpI++)
            {
                Shape TShp = DraftDoc.Shapes[ShpI];

                TShp.ConvertToInlineShape().Select();

                WCSelection wcs = new WCSelection();

                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;

                DraftShapes.Add(wcs);
            }

            DraftIShapes.Sort((a, b) => a.ContentSelectionStart.CompareTo(b.ContentSelectionStart));
            DraftShapes.Sort((a, b) => a.ContentSelectionStart.CompareTo(b.ContentSelectionStart));

            WCSelection[] NewAArray = WCSelectionOperations.CreateNewArrangedSelectionArray(DraftIShapes.ToArray(), DraftShapes.ToArray());

            for(int ICount = 0; ICount < NewAArray.Length; ICount++)
            {
                NewAArray[ICount].ContentID = "image_" + (ICount + 1);
            }

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
            return NewAArray;
        }
    }
}
