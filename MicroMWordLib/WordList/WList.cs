using System;
using System.Collections.Generic;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordParagraph;
using MicroMWordLib.WordTable;
using MicroMWordLib.WordImage;
namespace MicroMWordLib.WordList
{
    public class WList : IWBaseElement
    {
        private string prp_ListName = null;
        private List<WListItem> prp_ListItems = null;
        private WCSelection prp_ContentSelection = null;

        public string ListName { get => prp_ListName; set => prp_ListName = value; }
        public List<WListItem> ListItems { get => prp_ListItems; set => prp_ListItems = value; }
        public WCSelection ContentSelection { get => prp_ContentSelection; set => prp_ContentSelection = value; }

        public WList()
        {
            prp_ListItems = new List<WListItem>();
        }

        public static WList RecoverInnerContentSelection(WList in_WList, int WCSelectionIndex)
        {
            bool updateWCSEnd = false;
            if(in_WList.ContentSelection == null)
            {
                in_WList.ContentSelection = new WCSelection() { ContentID = "WList_" + Guid.NewGuid(), ContentSelectionStart = WCSelectionIndex };
                updateWCSEnd = true;
            }

            for(int iindx = 0; iindx < in_WList.ListItems.Count; iindx++)
            {
                foreach(WListItemElement ItemEl in in_WList.ListItems[iindx].ItemElements)
                {
                    if (ItemEl.ListItemElement.GetType() == typeof(WParagraph))
                    {
                        ItemEl.ListItemElement = WParagraph.RecoverInnerContentSelection(ItemEl.ListItemElement as WParagraph, WCSelectionIndex);
                        WCSelectionIndex = (ItemEl.ListItemElement as WParagraph).ContentSelection.ContentSelectionEnd;
                    }
                    else
                    {
                        ItemEl.ListItemElement = WTable.RecoverInnerContentSelection(ItemEl.ListItemElement as WTable, WCSelectionIndex);
                        WCSelectionIndex = (ItemEl.ListItemElement as WTable).ContentSelection.ContentSelectionEnd;
                    }
                }

                if(in_WList.ListItems[iindx].ChildList != null)
                {
                    in_WList.ListItems[iindx].ChildList = RecoverInnerContentSelection(in_WList.ListItems[iindx].ChildList, WCSelectionIndex);

                    WCSelectionIndex = in_WList.ListItems[iindx].ChildList.ContentSelection.ContentSelectionEnd;
                }
            }

            if(updateWCSEnd == true)
            {
                in_WList.ContentSelection.ContentSelectionEnd = WCSelectionIndex;
            }

            return in_WList;
        }

        public int RecoverInnerContentSelection(WCSelection[] in_WCSelections)
        {
            bool updateWCSEnd = false;
            if (ContentSelection == null)
            {
                ContentSelection = new WCSelection() { ContentID = "WList_" + Guid.NewGuid(), ContentSelectionStart = in_WCSelections[0].ContentSelectionStart };
                updateWCSEnd = true;
            }

            int _SIndex = 0;

            for (int iindx = 0; iindx < ListItems.Count; iindx++)
            {

                foreach (WListItemElement ItemEl in ListItems[iindx].ItemElements)
                {
                    if (_SIndex > in_WCSelections.Length - 1)
                    {
                        return _SIndex;
                    }

                    if (ItemEl.ListItemElement.GetType() == typeof(WParagraph))
                    {
                        ItemEl.ListItemElement.ContentSelection = in_WCSelections[_SIndex];
                        _SIndex++;
                    }
                    else
                    {
                        WCSelection[] NSel = new WCSelection[in_WCSelections.Length - _SIndex];
                        Array.Copy(in_WCSelections, _SIndex, NSel, 0, NSel.Length);
                        _SIndex += (ItemEl.ListItemElement as WTable).RecoverInnerContentSelection(NSel);
                    }
                }

                if (ListItems[iindx].ChildList != null)
                {
                    WCSelection[] NSel = new WCSelection[in_WCSelections.Length - _SIndex];
                    Array.Copy(in_WCSelections, _SIndex, NSel, 0, NSel.Length);

                    _SIndex += ListItems[iindx].ChildList.RecoverInnerContentSelection(NSel);
                }
            }

            if (updateWCSEnd == true)
            {
                ContentSelection.ContentSelectionEnd = in_WCSelections[_SIndex - 1].ContentSelectionEnd;
            }

            return _SIndex;
        }

        public static WList RecoverImages(WList in_WList, WImage[] in_WImages)
        {
            if(in_WList.ContentSelection == null)
            {
                return in_WList;
            }

            WImage[] LImages;
            {
                List<WImage> ALLTImages = new List<WImage>();

                for (int img = 0; img < in_WImages.Length; img++)
                {
                    if(in_WImages[img].ContentSelection == null)
                    {
                        continue;
                    }

                    if (in_WImages[img].ContentSelection.ContentSelectionStart >= in_WList.ContentSelection.ContentSelectionStart && in_WImages[img].ContentSelection.ContentSelectionEnd <= in_WList.ContentSelection.ContentSelectionEnd)
                    {
                        ALLTImages.Add(in_WImages[img]);
                    }
                }
                LImages = ALLTImages.ToArray();
            }

            if(LImages.Length < 1)
            {
                return in_WList;
            }

            //int _LastSelection = -50;
            for (int lelem = 0; lelem < in_WList.ListItems.Count; lelem++)
            {
                for(int ielem = 0; ielem < in_WList.ListItems[lelem].ItemElements.Count; ielem++)
                {
                    if(in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement.GetType() == typeof(WParagraph))
                    {
                        //if (in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement.ContentSelection.ContentSelectionStart - _LastSelection > 20)
                        //{
                            in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement = WParagraph.RecoverImages(in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement as WParagraph, LImages);

                            //_LastSelection = in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement.ContentSelection.ContentSelectionEnd;
                        //}
                    }
                    else if(in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement.GetType() == typeof(WTable))
                    {
                        in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement = WTable.RecoverImages(in_WList.ListItems[lelem].ItemElements[ielem].ListItemElement as WTable, LImages);
                    }

                    if(in_WList.ListItems[lelem].ChildList != null)
                    {
                        in_WList.ListItems[lelem].ChildList = RecoverImages(in_WList.ListItems[lelem].ChildList, LImages);
                    }
                }
            }

            return in_WList;
        }

        public int RecoverImages(WImage[] in_WImages)
        {
            int _AICount = 0;

            WImage[] LImages;
            {
                List<WImage> ALLTImages = new List<WImage>();

                for (int img = 0; img < in_WImages.Length; img++)
                {
                    if (in_WImages[img].ContentSelection.ContentSelectionStart >= ContentSelection.ContentSelectionStart && in_WImages[img].ContentSelection.ContentSelectionEnd <= ContentSelection.ContentSelectionEnd)
                    {
                        ALLTImages.Add(in_WImages[img]);
                    }
                }
                LImages = ALLTImages.ToArray();
            }

            //int _LastSelection = -50;
            for (int lelem = 0; lelem < ListItems.Count; lelem++)
            {
                for (int ielem = 0; ielem < ListItems[lelem].ItemElements.Count; ielem++)
                {
                    if (ListItems[lelem].ItemElements[ielem].ListItemElement.GetType() == typeof(WParagraph))
                    {
                        //if (ListItems[lelem].ItemElements[ielem].ListItemElement.ContentSelection.ContentSelectionStart - _LastSelection > 20)
                        //{
                           _AICount += (ListItems[lelem].ItemElements[ielem].ListItemElement as WParagraph).RecoverImages(LImages);

                        //    _LastSelection = ListItems[lelem].ItemElements[ielem].ListItemElement.ContentSelection.ContentSelectionEnd;
                        //}
                    }
                    else if (ListItems[lelem].ItemElements[ielem].ListItemElement.GetType() == typeof(WTable))
                    {
                        _AICount += (ListItems[lelem].ItemElements[ielem].ListItemElement as WTable).RecoverImages(LImages);
                        //_LastSelection = -50;
                    }
                }
            }

            return _AICount;
        }

        public static WList[] RecoverImages(WList[] in_WLists, WImage[] in_WImages)
        {
            for(int wimg = 0; wimg < in_WLists.Length; wimg++)
            {

                if(in_WLists[wimg].ContentSelection == null)
                {
                    continue;
                }

                in_WLists[wimg] = RecoverImages(in_WLists[wimg], in_WImages);
            }

            return in_WLists;
        }

    }
}
