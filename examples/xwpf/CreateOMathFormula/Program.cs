using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateOMathFormula
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();
            var p = doc.CreateParagraph();
            var math = p.CreateOMath();

            var acc = math.CreateAcc();
            acc.AccPr = "¯";
            acc.Element.CreateRun().SetText("X");
            math.CreateRun().SetText("=");
            var f = math.CreateF();
            f.FractionType = ST_FType.bar;
            f.Denominator.CreateRun().SetText("n");
            var nary = f.Numerator.CreateNary().SetSumm();
            nary.Superscript.CreateRun().SetText("n");
            nary.Subscript.CreateRun().SetText("i=1");
            var ssub = nary.Element.CreateSSub();
            ssub.Element.CreateRun().SetText("X");
            ssub.Subscript.CreateRun().SetText("i");

            FileStream out1 = new FileStream("Xavg.docx", FileMode.Create);
            doc.Write(out1);
            out1.Close();
        }
    }
}
