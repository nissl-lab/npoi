using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.InsertPicturesInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("PictureSheet");


            IDrawing patriarch = sheet1.CreateDrawingPatriarch();
            //create the anchor
            XSSFClientAnchor anchor = new XSSFClientAnchor(500, 200, 0, 0, 2, 2, 4, 7);
            anchor.AnchorType =  AnchorType.MoveDontResize;
            //load the picture and get the picture index in the workbook
            //first picture
            int imageId= LoadImage("../../image/HumpbackWhale.jpg", workbook);
            XSSFPicture picture = (XSSFPicture)patriarch.CreatePicture(anchor, imageId);
            //Reset the image to the original size.
            //picture.Resize();   //Note: Resize will reset client anchor you set.
            picture.LineStyle = LineStyle.DashDotGel;

            //second picture
            int imageId2 = LoadImage("../../image/HumpbackWhale.jpg", workbook);
            XSSFClientAnchor anchor2 = new XSSFClientAnchor(500, 200, 0, 0, 5, 10, 7, 15);
            XSSFPicture picture2 = (XSSFPicture)patriarch.CreatePicture(anchor2, imageId2);
            picture.LineStyle = LineStyle.DashDotGel;

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }

        public static int LoadImage(string path, IWorkbook wb)
        {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            return wb.AddPicture(buffer, PictureType.JPEG);

        }
    }
}
