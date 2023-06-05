using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.Collections.Generic;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    internal class TestXSSFShape
    {

        [Test]
        public void TestShapeLineEndingCapType()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            XSSFClientAnchor anchor = new XSSFClientAnchor(0,0,0,0, 1,1,2,2);

            XSSFConnector cxn = drawing.CreateConnector(anchor);
            cxn.Name = "sp1";
            cxn.LineEndingCapType = NPOI.SS.UserModel.LineEndingCapType.Round;

            XSSFWorkbook rbwb = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);

            XSSFSheet rb_sht = (XSSFSheet)rbwb.GetSheetAt(0);
            XSSFDrawing dr = (XSSFDrawing)rb_sht.GetDrawingPatriarch();
            List<XSSFShape> lstShp = dr.GetShapes();
            foreach(var sp in lstShp)
            {
                if(sp.Name == "sp1")
                {
                    Assert.AreEqual(NPOI.SS.UserModel.LineEndingCapType.Round, sp.LineEndingCapType);
                    return;
                }
            }
            Assert.True(false);
        }
        [Test]
        public void TestShapeCompoundLineType()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            XSSFClientAnchor anchor = new XSSFClientAnchor(0,0,0,0, 1,1,2,2);

            XSSFConnector cxn = drawing.CreateConnector(anchor);
            cxn.Name = "sp2";
            cxn.CompoundLineType = NPOI.SS.UserModel.CompoundLineType.DoubleLines;

            XSSFWorkbook rbwb = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);

            XSSFSheet rb_sht = (XSSFSheet)rbwb.GetSheetAt(0);
            XSSFDrawing dr = (XSSFDrawing)rb_sht.GetDrawingPatriarch();
            List<XSSFShape> lstShp = dr.GetShapes();
            foreach(var sp in lstShp)
            {
                if(sp.Name == "sp2")
                {
                    Assert.AreEqual(NPOI.SS.UserModel.CompoundLineType.DoubleLines, sp.CompoundLineType);
                    return;
                }
            }
            Assert.True(false);
        }
    }
}
