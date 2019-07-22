using System.Collections.Generic;
namespace MicroMWordLib.WordContentSelection
{
    public abstract class WCSelectionOperations
    {
        public static WCSelection[] CreateNewArrangedSelectionArray(WCSelection[] in_List1, WCSelection[] in_List2)
        {
            List<WCSelection> DList1 = new List<WCSelection>(in_List1);
            List<WCSelection> DList2 = new List<WCSelection>(in_List2);
            DList1.Sort((a, b) => a.ContentSelectionStart.CompareTo(b.ContentSelectionStart));
            DList2.Sort((a, b) => a.ContentSelectionStart.CompareTo(b.ContentSelectionStart));
            in_List1 = DList1.ToArray();
            in_List2 = DList2.ToArray();

            List<WCSelection> out_ListArranged = new List<WCSelection>();

            out_ListArranged.Clear();

            int L2Index = 0;
            bool IndexOI = false;

            for (int List1 = 0; List1 < in_List1.Length; List1++)
            {
                if (L2Index < in_List2.Length)
                {
                    while (IndexOI == false)
                    {
                        if (in_List2[L2Index].ContentSelectionEnd < in_List1[List1].ContentSelectionStart)
                        {
                            out_ListArranged.Add(in_List2[L2Index]);
                            L2Index++;
                        }
                        else
                        {
                            out_ListArranged.Add(in_List1[List1]);
                            break;
                        }

                        if (L2Index >= in_List2.Length)
                        {
                            IndexOI = true;
                            out_ListArranged.Add(in_List1[List1]);
                        }

                    }
                }
                else
                {
                    out_ListArranged.Add(in_List1[List1]);
                }
            }

            for (int itm = L2Index; itm < in_List2.Length; itm++)
            {
                out_ListArranged.Add(in_List2[itm]);
            }

            return out_ListArranged.ToArray();
        }

        public static WCSelection[] RemoveCompairingParts(WCSelection[] MainWCSelection, WCSelection[] WCSelectionToCompare)
        {
            List<WCSelection> LWCSelectionMain = new List<WCSelection>(MainWCSelection);
            List<WCSelection> LWCSelectionToCompare = new List<WCSelection>(WCSelectionToCompare);

            int _mainLSCount = LWCSelectionMain.Count;
            int _cmprdLSelCount = LWCSelectionToCompare.Count;

            int _cmprdLSelIndex = 0;
            int _mainLSelIndex = 0;

            while (_mainLSelIndex < _mainLSCount)
            {
                if(_cmprdLSelIndex < _cmprdLSelCount)
                {
                    if (LWCSelectionMain[_mainLSelIndex].ContentSelectionStart >= LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart && LWCSelectionMain[_mainLSelIndex].ContentSelectionStart < LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd)
                    {

                        // First sub case
                        if (LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd <= LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd)
                        {
                            LWCSelectionMain.RemoveAt(_mainLSelIndex);
                            _mainLSCount = LWCSelectionMain.Count;
                        }
                        // Second sub case
                        else
                        {

                            LWCSelectionMain.Insert(_mainLSelIndex + 1, new WCSelection() { ContentID = null, ContentSelectionStart = LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd, ContentSelectionEnd = LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd });

                            LWCSelectionMain.RemoveAt(_mainLSelIndex);
                            _cmprdLSelIndex++;
                        }

                    }
                    else if (LWCSelectionMain[_mainLSelIndex].ContentSelectionStart < LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart)
                    {
                        // First sub case
                        if (LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd <= LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart)
                        {
                            _mainLSelIndex++;
                        }
                        // Second sub case
                        else if (LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd > LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart && LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd <= LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd)
                        {
                            LWCSelectionMain.Insert(_mainLSelIndex + 1, new WCSelection() { ContentID = null, ContentSelectionStart = LWCSelectionMain[_mainLSelIndex].ContentSelectionStart, ContentSelectionEnd = LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart });

                            LWCSelectionMain.RemoveAt(_mainLSelIndex);
                            _mainLSelIndex++;

                        }
                        else
                        {
                            LWCSelectionMain.Insert(_mainLSelIndex + 1, new WCSelection() { ContentID = null, ContentSelectionStart = LWCSelectionMain[_mainLSelIndex].ContentSelectionStart, ContentSelectionEnd = LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart });

                            LWCSelectionMain.Insert(_mainLSelIndex + 2, new WCSelection() { ContentID = null, ContentSelectionStart = LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd, ContentSelectionEnd = LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd });

                            LWCSelectionMain.RemoveAt(_mainLSelIndex);
                            _mainLSCount = LWCSelectionMain.Count;
                            _mainLSelIndex++;
                            _cmprdLSelIndex++;
                        }
                    }
                    else
                    {
                        _cmprdLSelIndex++;
                    }
                }
                else
                {
                    break;
                }
               
            }

            return LWCSelectionMain.ToArray();
        }

        public static WCSelection[] RemoveAdditonalCompairingParts(WCSelection[] MainWCSelection, WCSelection[] WCSelectionToCompare)
        {
            List<WCSelection> LWCSelectionMain = new List<WCSelection>(MainWCSelection);
            List<WCSelection> LWCSelectionToCompare = new List<WCSelection>(WCSelectionToCompare);

            int _mainLSCount = LWCSelectionMain.Count;
            int _cmprdLSelCount = LWCSelectionToCompare.Count;

            int _cmprdLSelIndex = 0;
            int _mainLSelIndex = 0;
            bool ComparinigFound = false;

            while (_cmprdLSelIndex < _cmprdLSelCount)
            {
                if (_mainLSelIndex < _mainLSCount)
                {
                    if (ComparinigFound == true)
                    {
                        if (LWCSelectionMain[_mainLSelIndex].ContentSelectionStart < LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd)
                        {
                            if (LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd > LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd)
                            {
                                LWCSelectionMain.Insert(_mainLSelIndex + 1, new WCSelection() { ContentID = null, ContentSelectionStart = LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd, ContentSelectionEnd = LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd });

                                LWCSelectionMain.RemoveAt(_mainLSelIndex);
                                _cmprdLSelIndex++;
                                ComparinigFound = false;
                            }
                            else if(LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd == LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionEnd)
                            {
                                LWCSelectionMain.RemoveAt(_mainLSelIndex);
                                _cmprdLSelIndex++;
                                ComparinigFound = false;
                            }
                            else
                            {
                                LWCSelectionMain.RemoveAt(_mainLSelIndex);
                            }
                        }
                    }
                    else
                    {
                        if (LWCSelectionMain[_mainLSelIndex].ContentSelectionStart <= LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart && LWCSelectionMain[_mainLSelIndex].ContentSelectionEnd > LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart)
                        {
                            ComparinigFound = true;
                            _mainLSelIndex++;
                        }
                        else if(LWCSelectionMain[_mainLSelIndex].ContentSelectionStart > LWCSelectionToCompare[_cmprdLSelIndex].ContentSelectionStart)
                        {
                            _cmprdLSelIndex++;
                        }
                        else
                        {
                            _mainLSelIndex++;
                        }
                    }
                }
                else
                {
                    break;
                }

            }
            return LWCSelectionMain.ToArray();
        }

        public static WCSelection[] RemoveChilds(WCSelection[] MainWCSelections, WCSelection[] WCSelectionsToCompare)
        {
            if (WCSelectionsToCompare == null)
            {
                return MainWCSelections;
            }
            
            if(MainWCSelections == null)
            {
                return null;
            }

            List<WCSelection> ListWCS = new List<WCSelection>(MainWCSelections);

            int _Lindex = 0;
            int _TIndex = 0;
            int _LCount = ListWCS.Count;
            int _LPos = 0;

            while (_Lindex < _LCount)
            {
                _TIndex = _LPos;
                while (_TIndex < WCSelectionsToCompare.Length)
                {
                    if (ListWCS[_Lindex].ContentSelectionStart >= WCSelectionsToCompare[_TIndex].ContentSelectionStart && ListWCS[_Lindex].ContentSelectionEnd <= WCSelectionsToCompare[_TIndex].ContentSelectionEnd)
                    {
                        ListWCS.RemoveAt(_Lindex);
                        _LCount = ListWCS.Count;
                        _LPos = _TIndex;
                        break;
                    }

                    _TIndex++;
                }

                _Lindex++;
            }
            return ListWCS.ToArray();
        }

        public static WCSelection[] RemoveChilds(WCSelection[] WCSelections)
        {
            if (WCSelections == null)
            {
                return null;
            }

            List<WCSelection> ListWCS = new List<WCSelection>(WCSelections);

            int _lsindex = 0;
            int _lscount = ListWCS.Count;
            int _csindex = 0;
            while (_csindex < _lscount)
            {
                _lsindex = 0;
                while (_lsindex < _lscount)
                {
                    if (ListWCS[_csindex].ContentSelectionStart >= ListWCS[_lsindex].ContentSelectionStart && ListWCS[_csindex].ContentSelectionEnd <= ListWCS[_lsindex].ContentSelectionEnd && _csindex != _lsindex)
                    {
                        ListWCS.RemoveAt(_csindex);
                        _lscount = ListWCS.Count;
                        _csindex--;
                        break;
                    }

                    _lsindex++;
                }

                _csindex++;
            }
            return ListWCS.ToArray();
        }

        public static WCSelection[] JoinSelections(WCSelection[] in_WSelections, int in_MinimumSelectionInterval, string in_SelectionID = null)
        {
            int LSIndex = 0;
            int ImgIndex = 0;

            List<WCSelection> NewASelection = new List<WCSelection>();

            while (LSIndex < in_WSelections.Length)
            {
                while (true)
                {
                    if (LSIndex + 1 < in_WSelections.Length)
                    {
                        if (in_WSelections[LSIndex + 1].ContentSelectionStart - in_WSelections[LSIndex].ContentSelectionEnd < in_MinimumSelectionInterval)
                        {
                            LSIndex++;
                        }
                        else break;
                    }
                    else break;
                }

                WCSelection wcs = new WCSelection();
                wcs.ContentSelectionStart = in_WSelections[ImgIndex].ContentSelectionStart;
                wcs.ContentSelectionEnd = in_WSelections[LSIndex].ContentSelectionEnd;
                NewASelection.Add(wcs);
                LSIndex++;
                ImgIndex = LSIndex;
            }

            if (in_SelectionID != null)
            {
                for (int iind = 0; iind < NewASelection.Count; iind++)
                {
                    NewASelection[iind].ContentID = in_SelectionID + (iind + 1);
                }
            }
            return NewASelection.ToArray();
        }
    }
}
