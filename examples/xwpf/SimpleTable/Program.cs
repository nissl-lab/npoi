using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XWPF.UserModel;
using System.IO;

namespace SimpleTable
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();

            XWPFTable table = doc.CreateTable(3, 3);

            table.GetRow(1).GetCell(1).SetText("EXAMPLE OF TABLE");


            XWPFParagraph p1 = doc.CreateParagraph();

            XWPFRun r1 = p1.CreateRun();
            r1.SetBold(true);
            r1.SetText("The quick brown fox");
            r1.SetBold(true);
            r1.SetFontFamily("Courier");
            r1.SetUnderline(UnderlinePatterns.DotDotDash);
            r1.SetTextPosition(100);

            table.GetRow(0).GetCell(0).SetParagraph(p1);


            table.GetRow(2).GetCell(2).SetText("only text");

            FileStream out1 = new FileStream("simpleTable.docx", FileMode.Create);
            doc.Write(out1);
            out1.Close();
        }
    }
}
