using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.Usermodel;
using NPOI.XWPF.UserModel;
using NUnit.Framework;using NUnit.Framework.Legacy;

namespace TestCases.XWPF.UserModel
{
    public class TestXWPFSharedRun
    {
        private CT_R ctRun;
        public IRunBody p;

        [SetUp]
        public void SetUp()
        {
            XWPFDocument doc = new XWPFDocument();
            p = doc.CreateParagraph();

            this.ctRun = new CT_R();
        }

        [Test]
        public void TestSetGetFontSize()
        {
            NPOI.OpenXmlFormats.Wordprocessing.CT_RPr rpr = ctRun.AddNewRPr1();
            rpr.AddNewSz().val = 14;

            XWPFSharedRun run = new XWPFSharedRun(ctRun, p);
            ClassicAssert.AreEqual(7.0, run.FontSize);

            run.FontSize = 24;
            ClassicAssert.AreEqual(48, (int) rpr.sz.val);

            run.FontSize = 24.5;
            ClassicAssert.AreEqual(24.5, run.FontSize);
        }
    }
}
