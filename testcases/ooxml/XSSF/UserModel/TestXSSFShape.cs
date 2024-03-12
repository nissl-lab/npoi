using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    internal class TestXSSFShapes
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
                switch(sp.Name)
                {
                    case "first":
                    case "second":
                    case "third":
                    case "L01":
                    case "L02":
                    case "L03":
                        break;
                    default:
                        Assert.Fail($"name is invalid [{sp.Name}]");
                        break;
                }
            }
        }

        [Test]
        public void TestShapeTextWrap()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("TestShapeTextWrap.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheet("Sheet1");
            XSSFDrawing drawing = sheet.GetDrawingPatriarch();

            List<XSSFShape> lstShp = drawing.GetShapes();
            foreach(var sp in lstShp)
            {
                if(sp.Name == "shape1")
                {
                    Assert.AreEqual(true, ((XSSFSimpleShape)sp).WordWrap);
                    XSSFTestDataSamples.WriteOut(wb, "TestShapeTextWrap-1-");
                    ((XSSFSimpleShape)sp).WordWrap = false;
                    XSSFWorkbook rbwb = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);
                    Assert.AreEqual(false, ((XSSFSimpleShape)sp).WordWrap);
                    XSSFTestDataSamples.WriteOut(wb, "TestShapeTextWrap-2-");
                }
            }
        }

        [Test]
        public void TestShapeInset()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sht0 = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sht0.CreateDrawingPatriarch();

            //----- create
            var ca0 = new XSSFClientAnchor(sht0, Units.ToEMU(100), Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200));
            var sp0 = drawing.CreateSimpleShape(ca0);
            sp0.LineStyle = LineStyle.Solid;
            sp0.LineStyleColor = 0xff00000;
            Assert.AreEqual(7.2, sp0.LeftInset);
            Assert.AreEqual(false, sp0.GetCTShape().txBody.bodyPr.IsSetLIns());
            Assert.AreEqual(3.6, sp0.TopInset);
            Assert.AreEqual(false, sp0.GetCTShape().txBody.bodyPr.IsSetTIns());
            Assert.AreEqual(7.2, sp0.RightInset);
            Assert.AreEqual(false, sp0.GetCTShape().txBody.bodyPr.IsSetRIns());
            Assert.AreEqual(3.6, sp0.BottomInset);
            Assert.AreEqual(false, sp0.GetCTShape().txBody.bodyPr.IsSetBIns());
            XSSFTestDataSamples.WriteOut(wb, "TestShapeTextWrap-1-");

            //----- zero
            sp0.LeftInset = 0;
            sp0.TopInset = 0;
            sp0.RightInset = 0;
            sp0.BottomInset = 0;
            XSSFWorkbook rbwb1 = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);
            var sht1 = rbwb1.GetSheet("sheet0");
            var dw1 = ((XSSFSheet)sht1).GetDrawingPatriarch();
            XSSFSimpleShape sp1 = (XSSFSimpleShape)dw1.GetShapes()[0];
            Assert.AreEqual(0, sp1.LeftInset);
            Assert.AreEqual(true, sp1.GetCTShape().txBody.bodyPr.IsSetLIns());
            Assert.AreEqual(0, sp1.TopInset);
            Assert.AreEqual(true, sp1.GetCTShape().txBody.bodyPr.IsSetTIns());
            Assert.AreEqual(0, sp1.RightInset);
            Assert.AreEqual(true, sp1.GetCTShape().txBody.bodyPr.IsSetRIns());
            Assert.AreEqual(0, sp1.BottomInset);
            Assert.AreEqual(true, sp1.GetCTShape().txBody.bodyPr.IsSetBIns());
            XSSFTestDataSamples.WriteOut(rbwb1, "TestShapeTextWrap-2-");

            //----- others
            sp1.LeftInset = 3.6;
            sp1.TopInset = 1.8;
            sp1.RightInset = 3.6;
            sp1.BottomInset = 1.8;
            XSSFWorkbook rbwb2 = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(rbwb1);
            var sht2 = rbwb2.GetSheet("sheet0");
            var dw2 = ((XSSFSheet)sht2).GetDrawingPatriarch();
            XSSFSimpleShape sp2 = (XSSFSimpleShape)dw2.GetShapes()[0];
            Assert.AreEqual(3.6, sp2.LeftInset);
            Assert.AreEqual(true, sp2.GetCTShape().txBody.bodyPr.IsSetLIns());
            Assert.AreEqual(1.8, sp2.TopInset);
            Assert.AreEqual(true, sp2.GetCTShape().txBody.bodyPr.IsSetTIns());
            Assert.AreEqual(3.6, sp2.RightInset);
            Assert.AreEqual(true, sp2.GetCTShape().txBody.bodyPr.IsSetRIns());
            Assert.AreEqual(1.8, sp2.BottomInset);
            Assert.AreEqual(true, sp2.GetCTShape().txBody.bodyPr.IsSetBIns());
            XSSFTestDataSamples.WriteOut(rbwb2, "TestShapeTextWrap-3-");
        }

        [Test]
        public void TestLockWithSheet()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            // simple shape
            XSSFClientAnchor ca0;
            ca0 = new XSSFClientAnchor(sheet, Units.ToEMU(100), Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200));
            XSSFSimpleShape shp0 = drawing.CreateSimpleShape(ca0);
            shp0.Name = "S00";
            shp0.cellanchor.clientData.fLocksWithSheet = true;

            // connector shape
            XSSFClientAnchor ca1;
            ca1 = new XSSFClientAnchor(sheet, Units.ToEMU(250), Units.ToEMU(250), Units.ToEMU(350), Units.ToEMU(350));
            XSSFConnector shp1 = drawing.CreateConnector(ca1);
            shp1.Name = "L00";
            shp1.cellanchor.clientData.fLocksWithSheet = true;

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
                        Assert.AreEqual(sp.cellanchor.clientData.fLocksWithSheet, true, "shape name:[{0}]", new object[] { sp.Name });
                        break;
                    default:
                        break;
                }
            }

            shp0.cellanchor.clientData.fLocksWithSheet = false;
            shp1.cellanchor.clientData.fLocksWithSheet = false;

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
                        Assert.AreEqual(sp.cellanchor.clientData.fLocksWithSheet, false, "shape name:[{0}]", new object[] { sp.Name });
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
            ganchor1 = new XSSFClientAnchor(sheet, Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200), Units.ToEMU(100));
            ganchor1.AnchorType = AnchorType.DontMoveAndResize;
            XSSFShapeGroup grp1= drawing.CreateGroup( ganchor1 );
            grp1.Name = "first";

            // connector shape
            XSSFChildAnchor ca1;
            ca1 = new XSSFChildAnchor(Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200), Units.ToEMU(100));
            XSSFConnector shp1 = grp1.CreateConnector( ca1 );
            shp1.Name = "L01";
            shp1.cellanchor.clientData.fLocksWithSheet = true;

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
            shp2.cellanchor.clientData.fLocksWithSheet = true;

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
            shp3.cellanchor.clientData.fLocksWithSheet = false;

            Assert.AreEqual(shp1.cellanchor.clientData.fLocksWithSheet, false);

            XSSFTestDataSamples.WriteOut(wb, "TestGroupLockWithSheet");
        }

        [Test]
        public void TestRecursiveGroup()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFFont font = wb.GetStylesSource().GetFontAt( 0 );
            font.SetScheme(FontScheme.NONE);

            //----- first group
            XSSFClientAnchor ganchor1;
            ganchor1 = new XSSFClientAnchor(sheet, Units.ToEMU(50), Units.ToEMU(400), Units.ToEMU(400), Units.ToEMU(50));
            ganchor1.AnchorType = AnchorType.DontMoveAndResize;
            XSSFShapeGroup grp1= drawing.CreateGroup( ganchor1 );
            grp1.Name = "G00";

            // simple shape
            XSSFChildAnchor ca1;
            ca1 = new XSSFChildAnchor(Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200), Units.ToEMU(100));
            XSSFSimpleShape shp1 = grp1.CreateSimpleShape( ca1 );
            shp1.Name = "S00";

            // connector shape
            XSSFChildAnchor ca2;
            ca2 = new XSSFChildAnchor(Units.ToEMU(100), Units.ToEMU(200), Units.ToEMU(200), Units.ToEMU(100));
            XSSFConnector shp2 = grp1.CreateConnector( ca2 );
            shp2.Name = "L00";
            shp2.cellanchor.clientData.fLocksWithSheet = true;

            int x = 300;
            int y = 300;
            int cx = 310;
            int cy = 310;
            XSSFShapeGroup grp = grp1;
            for(var ct = 1; ct < 5; ct++)
            {
                //----- second group
                XSSFChildGroupAnchor ganchor;
                ganchor = new XSSFChildGroupAnchor(Units.ToEMU(x - 25), Units.ToEMU(y - 25), Units.ToEMU(cx+25), Units.ToEMU(cy+25));
                grp = grp.CreateGroup(ganchor);
                grp.Name = $"G{ct}";

                // simple shape
                XSSFChildAnchor ca;
                ca = new XSSFChildAnchor(Units.ToEMU(x), Units.ToEMU(x), Units.ToEMU(cx), Units.ToEMU(cy));
                XSSFSimpleShape Sshp = grp.CreateSimpleShape( ca );
                Sshp.Name = $"S{ct:00}";
                Sshp.LineStyle = LineStyle.Solid;
                Sshp.LineStyleColor = 0x00FF00;

                // connector shape
                ca = new XSSFChildAnchor(Units.ToEMU(x), Units.ToEMU(y), Units.ToEMU(cx), Units.ToEMU(cy));
                XSSFConnector Cshp = grp.CreateConnector( ca );
                Cshp.Name = $"L{ct:00}";
                Cshp.cellanchor.clientData.fLocksWithSheet = true;

                x += 25;
                y += 25;
                cx += 25;
                cy += 25;
            }

            grp1.AutoFit(sheet);
            XSSFTestDataSamples.WriteOut(wb, "TestRecursiveGroup");
        }

        [Test]
        public void TestInnerPproduct()
        {
            DblVect2D b = new DblVect2D(1, 0);
            for(double Theta = 0; Theta < Math.PI; Theta+=Math.PI/180.0)
            {
                DblVect2D c = new DblVect2D(Math.Cos(Theta), Math.Sin(Theta));

                Debug.WriteLine($"{Theta/Math.PI*180}\t{Theta}\t{c.x}\t{c.y}\t={b.InnerProduct(c)}+$E$1");
            }
        }

        [Test]
        public void TestBuildFreeform_base()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFFont font = wb.GetStylesSource().GetFontAt( 0 );
            font.SetScheme(FontScheme.NONE);

            // free form shape
            var bff = new BuildFreeForm();
            bff.AddNode(new Coords(Units.ToEMU(100), Units.ToEMU(180)));
            bff.AddNode(new Coords(Units.ToEMU(200), Units.ToEMU(200)));
            bff.AddNode(new Coords(Units.ToEMU(300), Units.ToEMU(150)));
            bff.AddNode(new Coords(Units.ToEMU(400), Units.ToEMU(50)));
            if(bff.Build())
            {
                var shp = drawing.CreateFreeform(sheet, bff);
                shp.SetLineStyleColor(0, 0, 0);

                XSSFTestDataSamples.WriteOut(wb, "TestBuildFreeform_base");
            }
        }

        [Test]
        public void TestBuildFreeform()
        {
            XSSFWorkbook wb0 = XSSFTestDataSamples.OpenSampleWorkbook("TestBuildFreeform.xlsx");
            XSSFSheet sheet0 = (XSSFSheet)wb0.GetSheet("Sheet1");
            XSSFDrawing drawing0 = sheet0.GetDrawingPatriarch();
            List<XSSFShape> lstShp0 = drawing0.GetShapes();

            XSSFWorkbook wb1        = new XSSFWorkbook();
            XSSFSheet sheet1        = (XSSFSheet)wb1.CreateSheet(sheet0.SheetName);
            XSSFDrawing drawing1    = (XSSFDrawing)sheet1.CreateDrawingPatriarch();
            XSSFFont font           = wb1.GetStylesSource().GetFontAt( 0 );
            font.SetScheme(FontScheme.NONE);

            foreach(var sp in lstShp0)
            {
                if(sp.Name.Substring(0, 2) == "FF")
                //if(sp.Name.Substring(0, 4) == "FF01")
                {
                    var bff = new BuildFreeForm();

                    var spPr = ((XSSFSimpleShape)sp).GetCTShape().spPr;
                    var cg = spPr.custGeom;
                    for(int ct = 0; ct< cg.cxnLst.cxn.Count; ct++)
                    {
                        var cd = GetGeomGuide(cg.gdLst.gd, ct);
                        if( cd != null){
                            cd.Add(new Coords(spPr.xfrm.off.x, spPr.xfrm.off.y));
                            bff.AddNode(cd);
                        }
                    }
                    if(bff.Build())
                    {
                        var shp = drawing1.CreateFreeform(sheet1, bff);
                        shp.SetLineStyleColor(0, 0, 0);
                        shp.Name = sp.Name;
                    }
                }
            }
            XSSFTestDataSamples.WriteOut(wb1, "TestBuildFreeform");

            var rbwb = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb1);
            var rbsheet = (XSSFSheet) rbwb.GetSheet("Sheet1");
            var rbdrawing = rbsheet.GetDrawingPatriarch();
            var rblstShp = rbdrawing.GetShapes();
            foreach(var sp0 in lstShp0)
            {
                if(sp0.Name.Substring(0, 2) == "FF")
                {
                    var spPr0 = ((XSSFSimpleShape)sp0).GetCTShape().spPr;
                    var cg0 = spPr0.custGeom;

                    foreach(var rbsp in rblstShp)
                    {
                        if(sp0.Name == rbsp.Name)
                        {
                            var rbspPr = ((XSSFSimpleShape)rbsp).GetCTShape().spPr;
                            var rbcg = rbspPr.custGeom; //read back custom geometry

                            Assert.IsFalse(!Compeare(spPr0.xfrm.off.x, rbspPr.xfrm.off.x), $"{sp0.Name}-xfrm.off.x:{spPr0.xfrm.off.x}!={rbspPr.xfrm.off.x}");
                            Assert.IsFalse(!Compeare(spPr0.xfrm.off.y, rbspPr.xfrm.off.y), $"{sp0.Name}-xfrm.off.y:{spPr0.xfrm.off.y}!={rbspPr.xfrm.off.y}");
                            Assert.IsFalse(!Compeare(spPr0.xfrm.ext.cx, rbspPr.xfrm.ext.cx), $"{sp0.Name}-xfrm.off.x:{spPr0.xfrm.ext.cx}!={rbspPr.xfrm.ext.cx}");
                            Assert.IsFalse(!Compeare(spPr0.xfrm.ext.cy, rbspPr.xfrm.ext.cy), $"{sp0.Name}-xfrm.off.y:{spPr0.xfrm.ext.cy}!={rbspPr.xfrm.ext.cy}");

                            Assert.AreEqual(cg0.cxnLst.cxn.Count, rbcg.cxnLst.cxn.Count, $"Count:{cg0.cxnLst.cxn.Count}!={rbcg.cxnLst.cxn.Count}");

                            for(int ct = 0; ct< cg0.cxnLst.cxn.Count; ct++)
                            {
                                var cd0 = GetGeomGuide(cg0.gdLst.gd, ct);
                                if(cd0 != null)
                                {
                                    var rbcd = GetGeomGuide(rbcg.gdLst.gd, ct);
                                    if(rbcd != null)
                                    {
                                        Assert.IsFalse(!Compeare(cd0.x, rbcd.x), $"{sp0.Name}-gdLst.x{ct}:{cd0.x}!={rbcd.x}");
                                        Assert.IsFalse(!Compeare(cd0.y, rbcd.y), $"{sp0.Name}-gdLst.y{ct}:{cd0.y}!={rbcd.y}");
                                    }
                                    else
                                    {
                                        throw new Exception();
                                    }
                                }
                                else
                                {
                                    throw new Exception();
                                }
                            }
                            var ogph = cg0.pathLst.path[0];
                            var rbph = rbcg.pathLst.path[0];
                            Assert.IsFalse(!Compeare(ogph.w, rbph.w), $"{sp0.Name}-path.w:{ogph.w}!={rbph.w}");
                            Assert.IsFalse(!Compeare(ogph.h, rbph.h), $"{sp0.Name}-path.h:{ogph.h}!={rbph.h}");

                            Assert.IsFalse(!Compeare(ogph.moveto.pt.x, rbph.moveto.pt.x), $"{sp0.Name}-moveto.x:{ogph.moveto.pt.x}!={rbph.moveto.pt.x}");
                            Assert.IsFalse(!Compeare(ogph.moveto.pt.y, rbph.moveto.pt.y), $"{sp0.Name}-moveto.y:{ogph.moveto.pt.y}!={rbph.moveto.pt.y}");

                            Assert.AreEqual(ogph.cubicBezTo.Count, rbph.cubicBezTo.Count, $"cubicBezTo.Count:{ogph.cubicBezTo.Count}!={rbph.cubicBezTo.Count}");

                            for(int i = 0; i<ogph.cubicBezTo.Count; i++)
                            {
                                Assert.IsFalse(!Compeare(ogph.cubicBezTo[i].pt[0].x, rbph.cubicBezTo[i].pt[0].x), $"{ogph.cubicBezTo[i].pt[0].x},{rbph.cubicBezTo[i].pt[0].x}");
                                Assert.IsFalse(!Compeare(ogph.cubicBezTo[i].pt[0].y, rbph.cubicBezTo[i].pt[0].y), $"{ogph.cubicBezTo[i].pt[0].y},{rbph.cubicBezTo[i].pt[0].y}");
                                Assert.IsFalse(!Compeare(ogph.cubicBezTo[i].pt[1].x, rbph.cubicBezTo[i].pt[1].x), $"{ogph.cubicBezTo[i].pt[1].x},{rbph.cubicBezTo[i].pt[1].x}");
                                Assert.IsFalse(!Compeare(ogph.cubicBezTo[i].pt[1].y, rbph.cubicBezTo[i].pt[1].y), $"{ogph.cubicBezTo[i].pt[1].y},{rbph.cubicBezTo[i].pt[1].y}");
                                Assert.IsFalse(!Compeare(ogph.cubicBezTo[i].pt[2].x, rbph.cubicBezTo[i].pt[2].x), $"{sp0.Name}-cubicBezTo[{i}].pt[2].x:{ogph.cubicBezTo[i].pt[2].x},{rbph.cubicBezTo[i].pt[2].x}");
                                Assert.IsFalse(!Compeare(ogph.cubicBezTo[i].pt[2].y, rbph.cubicBezTo[i].pt[2].y), $"{sp0.Name}-cubicBezTo[{i}].pt[2].y:{ogph.cubicBezTo[i].pt[2].y},{rbph.cubicBezTo[i].pt[2].y}");
                            }
                        }
                    }
                }
            }
        }

        private Coords GetGeomGuide(
              List<CT_GeomGuide> ggs
            , int ct
        ) {
            long x = 0;
            long y = 0;

            var gdX = Search(ggs, $"connsiteX{ct}");
            if(gdX != null)
            {
                string[] fmla = gdX.fmla.Split(new char[] { ' ' });
                x = long.Parse(fmla[1]);
            }
            var gdY = Search(ggs, $"connsiteY{ct}");
            if(gdY != null)
            {
                string[] fmla = gdY.fmla.Split(new char[] { ' ' });
                y = long.Parse(fmla[1]);
            }
            if(gdX == null || gdY ==null)
            {
                return null;
            }

            return new Coords(x, y);
        }

        private CT_GeomGuide Search(List<CT_GeomGuide> Gds, string name)
        {
            foreach(var gd in Gds)
            {
                if(gd.name == name)
                {
                    return gd;
                }
            }
            return null;
        }
        private bool Compeare(long B, long C, long a = 5)
        {
            if(B-a<=C && C<=B+a)
            {
                return true;
            }
            return false;
        }

        private bool Compeare(string B, string C, long a = 5)
        {
            long b = long.Parse(B);
            long c = long.Parse(C);
            if(b-a<=c && c<=b+a)
            {
                return true;
            }
            return false;
        }
    }
}
