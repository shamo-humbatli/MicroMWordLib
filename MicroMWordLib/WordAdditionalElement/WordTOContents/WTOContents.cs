using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordOperations;

namespace MicroMWordLib.WordAdditionalElement.WordTOContents
{
    public sealed class WTOContents
    {
        //private List<WAElementLine> prp_Lines;

        //public List<WAElementLine> Lines { get => prp_Lines; set => prp_Lines = value; }

        //public WTOContent()
        //{
        //    prp_Lines = new List<WAElementLine>();
        //}

        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument)
        {
            Document DraftDoc = MWordApp.Documents.Add();
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WCSelection> TOCSelection = new List<WCSelection>();

            for (int tcnt = 1; tcnt <= DraftDoc.TablesOfContents.Count; tcnt++)
            {
                TableOfContents toc = DraftDoc.TablesOfContents[tcnt];
               
                toc.Range.Select();
                WCSelection wcs = new WCSelection();
                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;
                TOCSelection.Add(wcs);
            }

            {
                for (int ls = 0; ls < TOCSelection.Count; ls++)
                {
                    TOCSelection[ls].ContentID = "WTOContents_" + (ls + 1);
                }
            }

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

            return TOCSelection.ToArray();
        }
    }
}
