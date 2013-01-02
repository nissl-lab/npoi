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
            anchor.AnchorType = 2;
            //load the picture and get the picture index in the workbook
            XSSFPicture picture = (XSSFPicture)patriarch.CreatePicture(anchor, LoadImage("../../image/HumpbackWhale.jpg", workbook));
            //Reset the image to the original size.
            //picture.Resize();   //Note: Resize will reset client anchor you set.
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
