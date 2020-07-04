using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;
using System.IO;

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

            InsertStandardDeviationFormula(doc.CreateParagraph());
            
            FileStream out1 = new FileStream("Xavg.docx", FileMode.Create);
            doc.Write(out1);
            out1.Close();
        }

        private static void InsertStandardDeviationFormula(XWPFParagraph p)
        {
            var math = p.CreateOMath();
            math.CreateRun().SetText("\u03c3=");

            var rad = math.CreateRad();
            rad.Degree.CreateRun().SetText("2");

            var f = rad.Element.CreateF();            
            f.FractionType = ST_FType.bar;
            f.Denominator.CreateRun().SetText("n-1");

            var nary = f.Numerator.CreateNary().SetSumm();
            nary.Superscript.CreateRun().SetText("n");
            nary.Subscript.CreateRun().SetText("i=1");

            var ssup = nary.Element.CreateSSup();
            ssup.Element.CreateRun().SetText("(");

            var ssub = ssup.Element.CreateSSub();
            ssub.Element.CreateRun().SetText("X");
            ssub.Subscript.CreateRun().SetText("i");
            ssup.Element.CreateRun().SetText("-");

            var acc = ssup.Element.CreateAcc();
            acc.AccPr = "¯";
            acc.Element.CreateRun().SetText("X");
            ssup.Element.CreateRun().SetText(")");
            ssup.Superscript.CreateRun().SetText("2");
        }
    }
}
