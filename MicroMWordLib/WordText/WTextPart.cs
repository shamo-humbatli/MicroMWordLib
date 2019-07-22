namespace MicroMWordLib.WordText
{
    public class WTextPart : IWParagraph
    {
        private bool prp_Bold = false;
        private bool prp_Italic = false;
        private bool prp_Underline = false;
        private string prp_Text = null;

        public string Text { get => prp_Text; set => prp_Text = value; }
        public bool Bold { get => prp_Bold; set => prp_Bold = value; }
        public bool Italic { get => prp_Italic; set => prp_Italic = value; }
        public bool Underline { get => prp_Underline; set => prp_Underline = value; }

        public WTextPart()
        {

        }

        public WTextPart(string PText)
        {
            prp_Text = PText;
        }
    }
}
