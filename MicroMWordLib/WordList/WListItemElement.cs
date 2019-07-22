namespace MicroMWordLib.WordList
{
    public class WListItemElement
    {
        private IWBaseElement prp_ListItemElement = null;
        private int prp_ListID = -1;
        private int prp_ListItemLevel = -1;

        public IWBaseElement ListItemElement { get => prp_ListItemElement; set => prp_ListItemElement = value; }
        public int ListID { get => prp_ListID; set => prp_ListID = value; }
        public int ListItemLevel { get => prp_ListItemLevel; set => prp_ListItemLevel = value; }

        public WListItemElement(IWBaseElement WElement, int WListID, int WListItemLevel)
        {
            prp_ListItemElement = WElement;
            prp_ListID = WListID;
            prp_ListItemLevel = WListItemLevel;
        }

        public WListItemElement() { }
    }
}
