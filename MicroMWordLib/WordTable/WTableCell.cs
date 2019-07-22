using System.Collections.Generic;

namespace MicroMWordLib.WordTable
{
    public enum TableCellValueType
    {
        Null = 0,
        Empty = 1,
        Text = 2
    }

    public enum TableCellType
    {
        Normal = 0,
        H_Merged = 1,
        V_Merged = 2,
        Deleted = 3
    }

    public class WTableCell
    {
        private TableCellType prp_WordTableCellType = TableCellType.Normal;

        //private TableCellValueType prp_WordTableCellValueType = TableCellValueType.Text;
        //private string prp_WordTableCellValue = string.Empty;
        private List<IWBaseElement> prp_CellElements = null;

        public TableCellType WordTableCellType { get => prp_WordTableCellType; set => prp_WordTableCellType = value; }
        public List<IWBaseElement> CellElements { get => prp_CellElements; set => prp_CellElements = value; }

        //public TableCellValueType WordTableCellValueType { get => prp_WordTableCellValueType; set => prp_WordTableCellValueType = value; }
        //public string WordTableCellValue { get => prp_WordTableCellValue; set => prp_WordTableCellValue = value; }
        //public List<WParagraph> CellParagraphs { get => prp_CellParagraphs; set => prp_CellParagraphs = value; }

        public WTableCell()
        {
            WordTableCellType = TableCellType.Normal;
            prp_CellElements = new List<IWBaseElement>();
            //this.WordTableCellValue = null;
            //this.WordTableCellValueType = TableCellValueType.Null;
        }

        public WTableCell(TableCellType _CellType = TableCellType.Normal)
        {
            WordTableCellType = _CellType;
            prp_CellElements = new List<IWBaseElement>();
            //this.WordTableCellValue = _CellValue;
            //this.WordTableCellValueType = _CellValueType;
        }


        //public string GetCellText(bool WithParagraphSeparator = true)
        //{
        //    if (this.Prp_CellParagraphs == null)
        //    {
        //        return null;
        //    }
        //    else if(this.Prp_CellParagraphs.Count < 1)
        //    {
        //        return string.Empty;
        //    }
        //    else
        //    {
        //        if (WithParagraphSeparator == true)
        //        {
        //            return string.Join(WTable.SeparatorFor_Paragraph, this.Prp_CellParagraphs.Select(ptxt => ptxt.GetText()).ToArray());
        //        }
        //        else
        //        {
        //            return string.Join(string.Empty, this.Prp_CellParagraphs.Select(ptxt => ptxt.GetText()).ToArray());
        //        }
        //    }
        //}
    }
}
