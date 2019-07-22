using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordOperations;

namespace MicroMWordLib.WordAdditionalElement.WordTOFigures
{
    public sealed class WTOFigures
    {
        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument)
        {
            Document DraftDoc = MWordApp.Documents.Add();
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WCSelection> TOCSelection = new List<WCSelection>();

            for (int tfgr = 1; tfgr <= DraftDoc.TablesOfFigures.Count; tfgr++)
            {
                TableOfFigures tof = DraftDoc.TablesOfFigures[tfgr];
                tof.Range.Select();
                WCSelection wcs = new WCSelection();
                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;
                TOCSelection.Add(wcs);
            }

            {
                for (int ls = 0; ls < TOCSelection.Count; ls++)
                {
                    TOCSelection[ls].ContentID = "WTOFigures_" + (ls + 1);
                }
            }

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

            return TOCSelection.ToArray();
        }
    }
}
