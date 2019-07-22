namespace MicroMWordLib.WordContentSelection
{
    public class WCSelection
    {
        private string prp_ContentID = null;
        private int prp_ContentSelectionStart = 0;
        private int prp_ContentSelectionEnd = 0;

        public string ContentID { get => prp_ContentID; set => prp_ContentID = value; }
        public int ContentSelectionStart { get => prp_ContentSelectionStart; set => prp_ContentSelectionStart = value; }
        public int ContentSelectionEnd { get => prp_ContentSelectionEnd; set => prp_ContentSelectionEnd = value; }
    }
}
