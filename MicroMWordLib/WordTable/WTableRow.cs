using System.Collections.Generic;

namespace MicroMWordLib.WordTable
{
    public class WTableRow
    {
        private List<WTableCell> prp_TableCells = null;
        private string prp_ConditionalFormatting = null;

        public List<WTableCell> TableCells { get => prp_TableCells; }
        public string ConditionalFormatting { get => prp_ConditionalFormatting; set => prp_ConditionalFormatting = value; }

        public WTableRow(WTableCell[] WTCells)
        {
            if (TableCells == null)
            {
                prp_TableCells = new List<WTableCell>();
            }

            TableCells.AddRange(WTCells);
        }

        public WTableRow()
        {
            if (prp_TableCells == null)
            {
                prp_TableCells = new List<WTableCell>();
            }
        }

        public void AddCell(WTableCell WTCell)
        {
            prp_TableCells.Add(WTCell);
        }

        public void AddCells(WTableCell[] WTCell)
        {
            prp_TableCells.AddRange(WTCell);
        }

    }
}
