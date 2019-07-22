using System.Collections.Generic;
namespace MicroMWordLib.WordList
{
    public class WListItem
    {
        private List<WListItemElement> prp_ItemElements = null;
        private WList prp_ChildList = null;
        private int prp_ListItemLevel = -1;
        private int prp_ListID = -1;

        public List<WListItemElement> ItemElements { get => prp_ItemElements; set => prp_ItemElements = value; }
        public WList ChildList { get => prp_ChildList; set => prp_ChildList = value; }
        public int ListItemLevel { get => prp_ListItemLevel; set => prp_ListItemLevel = value; }
        public int ListID { get => prp_ListID; set => prp_ListID = value; }

        public WListItem()
        {
            prp_ItemElements = new List<WListItemElement>();
        }
    }
}
