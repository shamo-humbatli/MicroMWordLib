namespace MicroMWordLib
{
    public abstract class WordXMLTags
    {
        private static string sntx_WordTagName_Body = "w:body";
        private static string sntx_WordTagName_Table = "w:tbl";
        private static string sntx_WordTagName_TableCell = "w:tc";
        private static string sntx_WordTagName_TableRow_Properties = "w:trPr";
        private static string sntx_WordTagName_TableCellProperty = "w:tcPr";
        private static string sntx_WordTagName_Paragraph = "w:p";
        private static string sntx_WordTagName_Paragraph_Properties = "w:pPr";
        private static string sntx_WordTagName_Paragraph_Properties_Style = "w:pStyle";
        private static string sntx_WordTagName_Paragraph_ListProperties = "w:listPr";
        private static string sntx_WordTagName_Paragraph_ListProperties_ListLevel1 = "w:ilvl";
        private static string sntx_WordTagName_Paragraph_ListProperties_ListLevel2 = "w:ilfo";
        private static string sntx_WordTagName_Paragraph_Properties_Alignment = "w:jc";
        private static string sntx_WordTagName_ConditionalFormatting = "w:cnfStyle";
        private static string sntx_WordTagName_Section = "wx:sect";
        private static string sntx_WordTagName_Hyperlink = "w:hlink";
        private static string sntx_WordTagName_Picture = "w:pict";



        private static string sntx_WordTagName_TextRun = "w:r";
        private static string sntx_WordTagName_Text = "w:t";

        private static string sntx_WordTagName_TableRow = "w:tr";
        private static string sntx_WordTagName_GridSpan = "w:gridSpan";
        private static string sntx_WordTagName_VerticalMerge = "w:vmerge";

        private static string sntx_WordTagAttrib_Value = "w:val";
        private static string sntx_WordTagAttrib_SectionAdditionRevisionID = "w:rsidR";

        private static string sntx_WordTagAttribValue_Restart = "restart";
        private static string sntx_WordTagAttribValue_Continue = "continue";
        private static string sntx_WordTagAttribValue_Center = "center";
        private static string sntx_WordTagAttribValue_Left = "left";
        private static string sntx_WordTagAttribValue_Right = "right";
        private static string sntx_WordTagAttribValue_Both = "both";
        private static string sntx_WordTagAttrib_Style = "style";

        private static string sntx_WordTagAttribValue_Heading1 = "Heading1";
        private static string sntx_WordTagAttribValue_Heading2 = "Heading2";
        private static string sntx_WordTagAttribValue_Heading3 = "Heading3";
        private static string sntx_WordTagAttribValue_Heading4 = "Heading4";
        private static string sntx_WordTagAttribValue_Heading5 = "Heading5";
        private static string sntx_WordTagAttribValue_Heading6 = "Heading6";
        private static string sntx_WordTagAttribValue_Heading7 = "Heading7";
        private static string sntx_WordTagAttribValue_Heading8 = "Heading8";
        private static string sntx_WordTagAttribValue_Heading9 = "Heading9";



        private static string sntx_WordTagName_TextRun_Properties = "w:rPr";
        private static string sntx_WordTagName_TextRun_Properties_Bold = "w:b";
        private static string sntx_WordTagName_TextRun_Properties_Italic = "w:i";
        private static string sntx_WordTagName_TextRun_Properties_Underline = "w:u";

        public static string WordTagName_TableCell { get => sntx_WordTagName_TableCell; }

        public static string WordTagName_TableCellProperty { get => sntx_WordTagName_TableCellProperty; }

        public static string WordTagName_Paragraph { get => sntx_WordTagName_Paragraph; }
        public static string WordTagName_TextRun { get => sntx_WordTagName_TextRun; }
        public static string WordTagName_Text { get => sntx_WordTagName_Text; }
        public static string WordTagName_TableRow { get => sntx_WordTagName_TableRow; }
        public static string WordTagName_GridSpan { get => sntx_WordTagName_GridSpan; }
        public static string WordTagName_VerticalMerge { get => sntx_WordTagName_VerticalMerge; }
        public static string WordTagAttrib_Value { get => sntx_WordTagAttrib_Value; }
        public static string WordTagAttribValue_Restart { get => sntx_WordTagAttribValue_Restart; }

        public static string WordTagAttribValue_Continue { get => sntx_WordTagAttribValue_Continue; }
        public static string WordTagName_TextRun_Properties_Underline { get => sntx_WordTagName_TextRun_Properties_Underline; }
        public static string WordTagName_TextRun_Properties_Italic { get => sntx_WordTagName_TextRun_Properties_Italic; }
        public static string WordTagName_TextRun_Properties_Bold { get => sntx_WordTagName_TextRun_Properties_Bold; }
        public static string WordTagName_TextRun_Properties { get => sntx_WordTagName_TextRun_Properties; }
        public static string WordTagAttribValue_Both { get => sntx_WordTagAttribValue_Both; }
        public static string WordTagAttribValue_Right { get => sntx_WordTagAttribValue_Right; }
        public static string WordTagAttribValue_Left { get => sntx_WordTagAttribValue_Left; }
        public static string WordTagAttribValue_Center { get => sntx_WordTagAttribValue_Center; }
        public static string WordTagName_Paragraph_Properties_Alignment { get => sntx_WordTagName_Paragraph_Properties_Alignment; }
        public static string WordTagName_Paragraph_Properties { get => sntx_WordTagName_Paragraph_Properties; }
        public static string WordTagName_ConditionalFormatting { get => sntx_WordTagName_ConditionalFormatting; }
        public static string WordTagName_TableRow_Properties { get => sntx_WordTagName_TableRow_Properties; }

        public static string WTN_Pgraph_LPrp { get => sntx_WordTagName_Paragraph_ListProperties; }
        public static string WTN_Pgraph_LPrp_LLvl1 { get => sntx_WordTagName_Paragraph_ListProperties_ListLevel1;  }
        public static string WTN_Pgraph_LPrp_LLvl2 { get => sntx_WordTagName_Paragraph_ListProperties_ListLevel2; }
        public static string WTN_Body { get => sntx_WordTagName_Body; }
        public static string WTN_Table { get => sntx_WordTagName_Table; }
        public static string WTN_Pgraph_Properties_Style { get => sntx_WordTagName_Paragraph_Properties_Style; }
        public static string WTA_SectionAdditionRevisionID { get => sntx_WordTagAttrib_SectionAdditionRevisionID; }
        public static string WTN_Section { get => sntx_WordTagName_Section; }
        public static string WTN_Hyperlink { get => sntx_WordTagName_Hyperlink; }
        public static string WTN_Picture { get => sntx_WordTagName_Picture; }
        public static string WTA_Style { get => sntx_WordTagAttrib_Style; }
    }
}
