using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XWPF.UserModel;
using System.IO;

namespace CreateEmptyDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();
            doc.CreateParagraph();
            FileStream sw = File.OpenWrite("blank.docx");
            doc.Write(sw);
            sw.Close();
        }
    }
}
