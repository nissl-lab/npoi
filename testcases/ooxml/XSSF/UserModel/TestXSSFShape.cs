using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;

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

        [Test]
        public void TestGetShapes()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("TestGetShapes.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheet("Sheet0");

            XSSFDrawing drawing = sheet.GetDrawingPatriarch();

            List<XSSFShape> lstShp = drawing.GetShapes();
            foreach(var sp in lstShp)
            {
                Debug.WriteLine($"name:[{sp.Name}]");
            }
        }

        [Test]
        public void TestLockWithSheet()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            // simple shape
            XSSFClientAnchor ca0;
            ca0 = sheet.CreateClientAnchor(Units.ToEMU(100), Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200));
            XSSFSimpleShape shp0 = drawing.CreateSimpleShape(ca0);
            shp0.Name = "S00";
            shp0.cttwocellanchor.clientData.fLocksWithSheet = true;

            // connector shape
            XSSFClientAnchor ca1;
            ca1 = sheet.CreateClientAnchor(Units.ToEMU(250), Units.ToEMU(250), Units.ToEMU(350), Units.ToEMU(350));
            XSSFConnector shp1 = drawing.CreateConnector(ca1);
            shp1.Name = "L00";
            shp1.cttwocellanchor.clientData.fLocksWithSheet = true;

            XSSFTestDataSamples.WriteOut(wb, "TestLockWithSheet1-");

            XSSFWorkbook rbwb = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);

            XSSFSheet rb_sht = (XSSFSheet)rbwb.GetSheetAt(0);
            XSSFDrawing dr = (XSSFDrawing)rb_sht.GetDrawingPatriarch();
            List<XSSFShape> lstShp = dr.GetShapes();
            foreach(var sp in lstShp)
            {
                switch(sp.Name)
                {
                    case "S00":
                    case "L00":
                        Assert.AreEqual(sp.cttwocellanchor.clientData.fLocksWithSheet, true, "shape name:[{0}]", new object[] { sp.Name });
                        break;
                    default:
                        break;
                }
            }

            shp0.cttwocellanchor.clientData.fLocksWithSheet = false;
            shp1.cttwocellanchor.clientData.fLocksWithSheet = false;

            XSSFTestDataSamples.WriteOut(wb, "TestLockWithSheet2-");

            XSSFWorkbook rbwb1 = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);
            XSSFSheet rb_sht1 = (XSSFSheet)rbwb1.GetSheetAt(0);
            XSSFDrawing dr1 = (XSSFDrawing)rb_sht1.GetDrawingPatriarch();
            List<XSSFShape> lstShp1 = dr1.GetShapes();

            foreach(var sp in lstShp1)
            {
                switch(sp.Name)
                {
                    case "S00":
                    case "L00":
                        Assert.AreEqual(sp.cttwocellanchor.clientData.fLocksWithSheet, false, "shape name:[{0}]", new object[] { sp.Name });
                        break;
                    default:
                        break;
                }
            }
        }

        [Test]
        public void TestGroupLockWithSheet()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            //----- first group
            XSSFClientAnchor ganchor1;
            ganchor1 = sheet.CreateClientAnchor(Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200), Units.ToEMU(100));
            ganchor1.AnchorType = AnchorType.DontMoveAndResize;
            XSSFShapeGroup grp1= drawing.CreateGroup( ganchor1 );
            grp1.Name = "first";

            // connector shape
            XSSFChildAnchor ca1;
            ca1 = new XSSFChildAnchor(Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200), Units.ToEMU(100));
            XSSFConnector shp1 = grp1.CreateConnector( ca1 );
            shp1.Name = "L01";
            shp1.cttwocellanchor.clientData.fLocksWithSheet = true;

            //----- second group
            XSSFChildGroupAnchor ganchor2;
            ganchor2 = new XSSFChildGroupAnchor(Units.ToEMU(110), Units.ToEMU(110), Units.ToEMU(190), Units.ToEMU(190));
            XSSFShapeGroup grp2 = grp1.CreateGroup( ganchor2 );
            grp2.Name = "second";

            // connector shape
            XSSFChildAnchor ca2;
            ca2 = new XSSFChildAnchor(Units.ToEMU(110), Units.ToEMU(110), Units.ToEMU(190), Units.ToEMU(190));
            XSSFConnector shp2 = grp2.CreateConnector(ca2);
            shp2.Name = "L02";
            shp2.cttwocellanchor.clientData.fLocksWithSheet = true;

            //----- third group
            XSSFChildGroupAnchor ganchor3;
            ganchor3 = new XSSFChildGroupAnchor(Units.ToEMU(120), Units.ToEMU(120), Units.ToEMU(180), Units.ToEMU(180));
            XSSFShapeGroup grp3 = grp2.CreateGroup( ganchor3 );
            grp3.Name = "third";

            // connector shape
            XSSFChildAnchor ca3;
            ca3 = new XSSFChildAnchor(Units.ToEMU(120), Units.ToEMU(150), Units.ToEMU(180), Units.ToEMU(150));
            XSSFConnector shp3 = grp3.CreateConnector(ca3);
            shp3.Name = "L03";
            shp3.cttwocellanchor.clientData.fLocksWithSheet = false;

            Assert.AreEqual(shp1.cttwocellanchor.clientData.fLocksWithSheet, false);

            XSSFTestDataSamples.WriteOut(wb, "TestGroupLockWithSheet");
        }
    }
}
