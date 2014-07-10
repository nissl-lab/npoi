using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace InsertPicturesInWord
{
    class Program
    {
        //http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
        //http://openxmltrix.blogspot.com/
        //http://stackoverflow.com/questions/7716078/formula-to-convert-net-pixels-to-excel-width-in-openxml-format


        const int emusPerInch = 914400;
        const int emusPerCm = 360000;

        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p2 = doc.CreateParagraph();
            XWPFRun r2 = p2.CreateRun();
            r2.SetText("test");


            var widthEmus = (int)(400.0 * 9525);
            var heightEmus = (int)(300.0 * 9525);

            using (FileStream picData = new FileStream("../../image/HumpbackWhale.jpg", FileMode.Open, FileAccess.Read))
            {
                r2.AddPicture(picData, (int)PictureType.PNG, "image1", widthEmus, heightEmus);
            }
            using (FileStream sw = File.Create("test.docx"))
            {
                doc.Write(sw);
            }
        }

    }
}
