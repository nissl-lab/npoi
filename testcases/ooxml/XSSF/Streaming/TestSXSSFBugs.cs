using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.Streaming;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCases.SS.UserModel;

namespace NPOI.OOXML.Testcases.XSSF.Streaming
{
    public class TestSXSSFBugs : BaseTestBugzillaIssues
    {
        public TestSXSSFBugs()
            : base(SXSSFITestDataProvider.instance)
        {

        }

        protected override void TrackColumnsForAutoSizingIfSXSSF(ISheet sheet)
        {
            SXSSFSheet sxSheet = (SXSSFSheet)sheet;
            sxSheet.TrackAllColumnsForAutoSizing();
        }
        [Test]
        public void Tug49253()
        {
            IWorkbook wb1 = new SXSSFWorkbook();
            IWorkbook wb2 = new SXSSFWorkbook();

            // No print settings before repeating
            ISheet s1 = wb1.CreateSheet();

            wb1.SetRepeatingRowsAndColumns(0, 2, 3, 1, 2);

            IPrintSetup ps1 = s1.PrintSetup;
            Assert.AreEqual(false, ps1.ValidSettings);
            Assert.AreEqual(false, ps1.Landscape);


            // Had valid print settings before repeating
            ISheet s2 = wb2.CreateSheet();
            IPrintSetup ps2 = s2.PrintSetup;

            ps2.Landscape = (false);
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);

            wb2.SetRepeatingRowsAndColumns(0, 2, 3, 1, 2);

            ps2 = s2.PrintSetup;
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);

            wb1.Close();
            wb2.Close();
        }
    }
}
