using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordOperations;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
namespace MicroMWordLib.WordImage
{
    public class WImageExporter
    {
        public static WImage[] ExportImages(Application MWordApp, Document MWordDocument, string OutputFolder, string ImageFileName)
        {
            WCSelection[] ImageWCSList;
            ImageWCSList = WImage.GetAllContentSelections(MWordApp, MWordDocument);
            List<WImage> ImageList = new List<WImage>();

            for (int wcsl = 0; wcsl < ImageWCSList.Length; wcsl++)
            {
                WImage wimg = new WImage();
                wimg.ImagePath = OutputFolder + "\\" + ImageFileName + "_" + ImageWCSList[wcsl].ContentID + ".png";
                wimg.ContentSelection = ImageWCSList[wcsl];
                ImageList.Add(wimg);
            }

            Document DraftDoc = MWordApp.Documents.Add(Visible: false);
            //Document DraftDoc = WordApp.Documents.Add(WParameters.Missing, WParameters.Missing, WParameters.Missing, false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            {
                DirectoryInfo OutpFol = new DirectoryInfo(OutputFolder);
                if (OutpFol.Exists == false)
                {
                    OutpFol.Create();
                }

                for (int isel = 0; isel < ImageList.Count; isel++)
                {

                    //MWordDocument.Range(ImageList[isel].ContentSelection.ContentSelectionStart, ImageList[isel].ContentSelection.ContentSelectionEnd).Select();

                    MWordApp.Selection.Start = ImageList[isel].ContentSelection.ContentSelectionStart;
                    MWordApp.Selection.End = ImageList[isel].ContentSelection.ContentSelectionEnd;

                    byte[] ImgData = MWordApp.Selection.Range.EnhMetaFileBits;
                    MemoryStream TMStream = new MemoryStream(ImgData);
                    Image ImgFS = Image.FromStream(TMStream);
                    ImgFS.Save(ImageList[isel].ImagePath, ImageFormat.Png);
                    ImgFS.Dispose();
                }
            }

            string[] ImgPaths = ImageList.Select(a => a.ImagePath).ToArray();
            ImageOperations.RemoveTransparentPartsFromImages(ImgPaths);

            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
            return ImageList.ToArray();
        }

        public static WImage[] ExportImages(Application MWordApp, Document MWordDocument, WCSelection[] WImageSelections, string OutputFolder, string ImageFileName)
        {

            List<WImage> ImageList = new List<WImage>();

            for (int wcsl = 0; wcsl < WImageSelections.Length; wcsl++)
            {
                WImage wimg = new WImage();
                wimg.ImagePath = OutputFolder + "\\" + ImageFileName + "_" + WImageSelections[wcsl].ContentID + ".jpg";
                wimg.ContentSelection = WImageSelections[wcsl];
                ImageList.Add(wimg);
            }

            Document DraftDoc = MWordApp.Documents.Add(Visible: false);
            //Document DraftDoc = WordApp.Documents.Add(WParameters.Missing, WParameters.Missing, WParameters.Missing, false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            {
                DirectoryInfo OutpFol = new DirectoryInfo(OutputFolder);
                if (OutpFol.Exists == false)
                {
                    OutpFol.Create();
                }

                for (int isel = 0; isel < ImageList.Count; isel++)
                {

                    //MWordDocument.Range(ImageList[isel].ContentSelection.ContentSelectionStart, ImageList[isel].ContentSelection.ContentSelectionEnd).Select();

                    MWordApp.Selection.Start = ImageList[isel].ContentSelection.ContentSelectionStart;
                    MWordApp.Selection.End = ImageList[isel].ContentSelection.ContentSelectionEnd;

                    byte[] ImgData = MWordApp.Selection.Range.EnhMetaFileBits;
                    MemoryStream TMStream = new MemoryStream(ImgData);
                    Image ImgFS = Image.FromStream(TMStream);
                    ImgFS.Save(ImageList[isel].ImagePath, ImageFormat.Png);
                    ImgFS.Dispose();
                }
            }

            string[] ImgPaths = ImageList.Select(a => a.ImagePath).ToArray();
            ImageOperations.RemoveTransparentPartsFromImages(ImgPaths);
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
            return ImageList.ToArray();
        }
    }
}
