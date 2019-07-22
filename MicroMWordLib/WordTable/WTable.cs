using System;
using System.Collections.Generic;
using System.Linq;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordParagraph;
using MicroMWordLib.WordList;
using MicroMWordLib.WordImage;
namespace MicroMWordLib.WordTable
{
    public class WTable : IWBaseElement
    {
        private string prp_TableName = string.Empty;
        private WCSelection prp_ContentSelection = null;
        private List<WTableRow> prp_TableRows = null;

        public WCSelection ContentSelection { get => prp_ContentSelection; set => prp_ContentSelection = value; }
        public List<WTableRow> TableRows { get => prp_TableRows; }
        public string TableName { get => prp_TableName; set => prp_TableName = value; }

        public WTable()
        {
            prp_TableRows = new List<WTableRow>();
        }

        public WTable(WTableRow[] WTRows)
        {
            prp_TableRows = new List<WTableRow>();

            prp_TableRows.AddRange(WTRows);
        }

        public void ArrangeTableCells()
        {
            if (prp_TableRows.Count > 1)
            {
                int maxcelll = prp_TableRows.Aggregate((c1, c2) => c1.TableCells.Count > c2.TableCells.Count ? c1 : c2).TableCells.Count;

                for (int r = 0; r < prp_TableRows.Count - 1; r++)
                {
                    int cc = TableRows[r].TableCells.Count;
                    while (cc < maxcelll)
                    {
                        prp_TableRows[r].AddCell(new WTableCell(TableCellType.Deleted));
                        cc = TableRows[r].TableCells.Count;
                    }
                }
            }
        }

        public void AddRow(WTableRow WTRow)
        {
            if (WTRow != null)
            {
                prp_TableRows.Add(WTRow);
            }
            ArrangeTableCells();
        }

        public void AddRowRange(WTableRow[] WTRowRange)
        {
            foreach (WTableRow wtr in WTRowRange)
            {
                if (wtr != null)
                {
                    prp_TableRows.Add(wtr);
                }
            }
            ArrangeTableCells();
        }

        public bool RemoveRow(int RowIndex)
        {
            if(RowIndex < 0 | RowIndex > prp_TableRows.Count - 1 | prp_TableRows.Count < 1)
            {
                return false;
            }

            this.prp_TableRows.RemoveAt(RowIndex);

            return true;
        }

        public WTableCell GetCell(int WTRowIndex, int WTColumnIndex)
        {
            if(WTRowIndex < 0 | WTColumnIndex < 0)
            {
                return null;
            }

            if(WTRowIndex >= prp_TableRows.Count)
            {
                return null;
            }

            if(WTColumnIndex >= TableRows[WTRowIndex].TableCells.Count)
            {
                return null;
            }

            return TableRows[WTRowIndex].TableCells[WTColumnIndex];
        }

        public bool ReplaceCell(WTableCell TableCell,  int WTRowIndex, int WTColumnIndex)
        {
            if (WTRowIndex < 0 | WTColumnIndex < 0)
            {
                return false;
            }

            if (WTRowIndex >= prp_TableRows.Count)
            {
                return false;
            }

            if (WTColumnIndex >= TableRows[WTRowIndex].TableCells.Count)
            {
                return false;
            }

            TableRows[WTRowIndex].TableCells[WTColumnIndex] = TableCell;
            return true;
        }

        public static WTable RecoverInnerContentSelection(WTable in_WTable, int WCSelectionIndex)
        {

            bool updateWCSEnd = false;
            if (in_WTable.ContentSelection == null)
            {
                in_WTable.ContentSelection = new WCSelection() { ContentID = "WTable_" + Guid.NewGuid(), ContentSelectionStart = WCSelectionIndex };
                updateWCSEnd = true;
            }

            for(int wtrow = 0; wtrow < in_WTable.TableRows.Count; wtrow++)
            {
                for(int wtcol = 0; wtcol < in_WTable.TableRows[wtrow].TableCells.Count; wtcol++)
                {
                    int _ElCount = 0;

                    for (int CElemIndx = 0; CElemIndx < in_WTable.TableRows[wtrow].TableCells[wtcol].CellElements.Count; CElemIndx++)
                    {
                        IWBaseElement CElem = in_WTable.TableRows[wtrow].TableCells[wtcol].CellElements[CElemIndx];
                        if (CElem.GetType() == typeof(WParagraph))
                        {
                            CElem = WParagraph.RecoverInnerContentSelection(CElem as WParagraph, WCSelectionIndex);
                            WCSelectionIndex = (CElem as WParagraph).ContentSelection.ContentSelectionEnd;
                        }
                        else if(CElem.GetType() == typeof(WList))
                        {
                            CElem = WList.RecoverInnerContentSelection(CElem as WList, WCSelectionIndex);
                            WCSelectionIndex = (CElem as WList).ContentSelection.ContentSelectionEnd;
                        }
                        else
                        {
                            CElem = RecoverInnerContentSelection(CElem as WTable, WCSelectionIndex);
                            WCSelectionIndex = (CElem as WTable).ContentSelection.ContentSelectionEnd;
                        }

                        _ElCount++;
                    }
                }
            }

            if (updateWCSEnd == true)
            {
                in_WTable.ContentSelection.ContentSelectionEnd = WCSelectionIndex;
            }


            return in_WTable;
        }

        public int RecoverInnerContentSelection(WCSelection[] in_WCSelections)
        {

            bool updateWCSEnd = false;
            if (ContentSelection == null)
            {
                ContentSelection = new WCSelection() { ContentID = "WTable_" + Guid.NewGuid(), ContentSelectionStart = in_WCSelections[0].ContentSelectionStart };
                updateWCSEnd = true;
            }

            int _SIndex = 0;
            int ls = 0;

            for (int wtrow = 0; wtrow < TableRows.Count; wtrow++)
            {
                for (int wtcol = 0; wtcol < TableRows[wtrow].TableCells.Count; wtcol++)
                {
                    for (int CElemIndx = 0; CElemIndx < TableRows[wtrow].TableCells[wtcol].CellElements.Count; CElemIndx++)
                    {
                        if (_SIndex > in_WCSelections.Length - 1)
                        {
                            return _SIndex;
                        }

                        IWBaseElement CElem = TableRows[wtrow].TableCells[wtcol].CellElements[CElemIndx];
                        if (CElem.GetType() == typeof(WParagraph))
                        {
                            CElem.ContentSelection = in_WCSelections[_SIndex];
                            _SIndex++;
                        }
                        else if (CElem.GetType() == typeof(WList))
                        {
                            WCSelection[] NSel = new WCSelection[in_WCSelections.Length - _SIndex];
                            Array.Copy(in_WCSelections, _SIndex, NSel, 0, NSel.Length);
                            _SIndex += (CElem as WList).RecoverInnerContentSelection(NSel);
                        }
                        else
                        {
                            WCSelection[] NSel = new WCSelection[in_WCSelections.Length - _SIndex];
                            Array.Copy(in_WCSelections, _SIndex, NSel, 0, NSel.Length);
                            _SIndex += (CElem as WTable).RecoverInnerContentSelection(NSel);
                        }
                    }
                }
            }

            if (updateWCSEnd == true)
            {
                ContentSelection.ContentSelectionEnd = in_WCSelections[_SIndex - 1].ContentSelectionEnd;
            }
            
            return _SIndex;
        }

        public static WTable RecoverImages(WTable in_WTable, WImage[] in_WImages)
        {

            if(in_WTable.ContentSelection == null)
            {
                return in_WTable;
            }

            WImage[] TImages;
            {
                List<WImage> ALLTImages = new List<WImage>();

                for (int img = 0; img < in_WImages.Length; img++)
                {
                    if (in_WImages[img].ContentSelection == null)
                    {
                        continue;
                    }

                    if (in_WImages[img].ContentSelection.ContentSelectionStart >= in_WTable.ContentSelection.ContentSelectionStart && in_WImages[img].ContentSelection.ContentSelectionEnd <= in_WTable.ContentSelection.ContentSelectionEnd)
                    {
                        ALLTImages.Add(in_WImages[img]);
                    }
                }
                TImages = ALLTImages.ToArray();
            }

            if(TImages.Length < 1)
            {
                return in_WTable;
            }

            for(int wtrow = 0; wtrow < in_WTable.TableRows.Count; wtrow++)
            {
                for(int wtcol = 0; wtcol < in_WTable.TableRows[wtrow].TableCells.Count; wtcol++)
                {
                    WTableCell WTCell = in_WTable.GetCell(wtrow, wtcol);
                    //int _LastSelection = -50;
                    for(int cel = 0; cel < WTCell.CellElements.Count; cel++)
                    {
                        IWBaseElement WCElement = WTCell.CellElements[cel];

                        if(WCElement.GetType() == typeof(WParagraph))
                        {
                            //if (WCElement.ContentSelection.ContentSelectionStart - _LastSelection > 20)
                            //{
                                WTCell.CellElements[cel] = WParagraph.RecoverImages(WTCell.CellElements[cel] as WParagraph, TImages);
                            //    _LastSelection = WCElement.ContentSelection.ContentSelectionEnd;
                            //}
                        }
                        else if(WCElement.GetType() == typeof(WList))
                        {
                            WTCell.CellElements[cel] = WList.RecoverImages(WTCell.CellElements[cel] as WList, TImages);
                        }
                        else if(WCElement.GetType() == typeof(WTable))
                        {
                            WTCell.CellElements[cel] = RecoverImages(WTCell.CellElements[cel] as WTable, TImages);
                        }
                    }
                }
            }

            return in_WTable;
        }

        public int RecoverImages(WImage[] in_WImages)
        {
           int _AICount = 0;
            WImage[] TImages;
            {
                List<WImage> ALLTImages = new List<WImage>();

                for (int img = 0; img < in_WImages.Length; img++)
                {
                    if (in_WImages[img].ContentSelection.ContentSelectionStart >= ContentSelection.ContentSelectionStart && in_WImages[img].ContentSelection.ContentSelectionEnd <= ContentSelection.ContentSelectionEnd)
                    {
                        ALLTImages.Add(in_WImages[img]);
                    }
                }
                TImages = ALLTImages.ToArray();
            }

            for (int wtrow = 0; wtrow < TableRows.Count; wtrow++)
            {
                for (int wtcol = 0; wtcol < TableRows[wtrow].TableCells.Count; wtcol++)
                {
                    WTableCell WTCell = GetCell(wtrow, wtcol);
                    //int _LastSelection = -50;
                    for (int cel = 0; cel < WTCell.CellElements.Count; cel++)
                    {
                        IWBaseElement WCElement = WTCell.CellElements[cel];

                        if (WCElement.GetType() == typeof(WParagraph))
                        {
                            //if (WCElement.ContentSelection.ContentSelectionStart - _LastSelection > 20)
                            //{
                                _AICount += (WTCell.CellElements[cel] as WParagraph).RecoverImages(TImages);
                            //    _LastSelection = WCElement.ContentSelection.ContentSelectionEnd;
                            //}
                        }
                        else if (WCElement.GetType() == typeof(WList))
                        {
                            WTCell.CellElements[cel] = WList.RecoverImages(WTCell.CellElements[cel] as WList, TImages);
                            //_LastSelection = -50;
                        }
                        else if (WCElement.GetType() == typeof(WTable))
                        {
                            _AICount += (WTCell.CellElements[cel] as WTable).RecoverImages(TImages);
                            //_LastSelection = -50;
                        }
                    }
                }
            }

            return _AICount;
        }

        public static WTable[] RecoverImages(WTable[] in_WTables, WImage[] in_WImages)
        {
            for (int wtbl = 0; wtbl < in_WTables.Length; wtbl++)
            {
                in_WTables[wtbl] = RecoverImages(in_WTables[wtbl], in_WImages);
            }

            return in_WTables;
        }
    }
}
