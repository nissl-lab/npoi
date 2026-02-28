using NPOI.XSSF;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.IO;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestCloneSheet
    {
        [Test]
        public void TestCloneHyperlink()
        {
            using(var workbook = XSSFTestDataSamples.OpenSampleWorkbook("1370_clonesheet_withhyperlink.xlsx"))
            {
                workbook.CloneSheet(0,"Sheet2");
            }
        }
    }
}