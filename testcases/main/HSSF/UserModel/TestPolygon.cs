using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Model;
using NPOI.DDF;
using TestCases.HSSF.Model;
using NPOI.Util;
using NPOI.HSSF.Record;

namespace TestCases.HSSF.UserModel
{

    /**
     * @author Evgeniy Berlog
     * @date 28.06.12
     */
    [TestFixture]
    public class TestPolygon
    {
        [Test]
        public void TestResultEqualsToAbstractShape()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(100, 100);
            polygon.SetPoints(new int[] { 0, 90, 50 }, new int[] { 5, 5, 44 });
            PolygonShape polygonShape = HSSFTestModelHelper.CreatePolygonShape(1024, polygon);
            polygon.ShapeId = (1024);

            Assert.AreEqual(polygon.GetEscherContainer().ChildRecords.Count, 4);
            Assert.AreEqual(polygonShape.SpContainer.ChildRecords.Count, 4);

            //sp record
            byte[] expected = polygonShape.SpContainer.GetChild(0).Serialize();
            byte[] actual = polygon.GetEscherContainer().GetChild(0).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = polygonShape.SpContainer.GetChild(2).Serialize();
            actual = polygon.GetEscherContainer().GetChild(2).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = polygonShape.SpContainer.GetChild(3).Serialize();
            actual = polygon.GetEscherContainer().GetChild(3).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            ObjRecord obj = polygon.GetObjRecord();
            ObjRecord objShape = polygonShape.ObjRecord;

            expected = obj.Serialize();
            actual = objShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));
        }
        [Test]
        public void TestPolygonPoints()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(100, 100);
            polygon.SetPoints(new int[] { 0, 90, 50, 90 }, new int[] { 5, 5, 44, 88 });

            PolygonShape polygonShape = HSSFTestModelHelper.CreatePolygonShape(0, polygon);

            EscherArrayProperty verticesProp1 = polygon.GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES) as EscherArrayProperty;
            EscherArrayProperty verticesProp2 = ((EscherOptRecord)polygonShape.SpContainer.GetChildById(EscherOptRecord.RECORD_ID))
                    .Lookup(EscherProperties.GEOMETRY__VERTICES) as EscherArrayProperty;

            Assert.AreEqual(verticesProp1.NumberOfElementsInArray, verticesProp2.NumberOfElementsInArray);
            Assert.AreEqual(verticesProp1.ToXml(""), verticesProp2.ToXml(""));

            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });
            Assert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));

            polygonShape = HSSFTestModelHelper.CreatePolygonShape(0, polygon);
            verticesProp1 = polygon.GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES) as EscherArrayProperty;
            verticesProp2 = ((EscherOptRecord)polygonShape.SpContainer.GetChildById(EscherOptRecord.RECORD_ID))
                    .Lookup(EscherProperties.GEOMETRY__VERTICES) as EscherArrayProperty;

            Assert.AreEqual(verticesProp1.NumberOfElementsInArray, verticesProp2.NumberOfElementsInArray);
            Assert.AreEqual(verticesProp1.ToXml(""), verticesProp2.ToXml(""));
        }
        [Test]
        public void TestSetGetProperties()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            Assert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));
            Assert.AreEqual(polygon.DrawAreaHeight, 101);
            Assert.AreEqual(polygon.DrawAreaWidth, 102);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            polygon = (HSSFPolygon)patriarch.Children[0];
            Assert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));
            Assert.AreEqual(polygon.DrawAreaHeight, 101);
            Assert.AreEqual(polygon.DrawAreaWidth, 102);

            polygon.SetPolygonDrawArea(1021, 1011);
            polygon.SetPoints(new int[] { 11, 21, 31 }, new int[] { 41, 51, 61 });

            Assert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 11, 21, 31 }));
            Assert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 41, 51, 61 }));
            Assert.AreEqual(polygon.DrawAreaHeight, 1011);
            Assert.AreEqual(polygon.DrawAreaWidth, 1021);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            polygon = (HSSFPolygon)patriarch.Children[0];

            Assert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 11, 21, 31 }));
            Assert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 41, 51, 61 }));
            Assert.AreEqual(polygon.DrawAreaHeight, 1011);
            Assert.AreEqual(polygon.DrawAreaWidth, 1021);
        }
        [Test]
        public void TestAddToExistingFile()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            HSSFPolygon polygon1 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon1.SetPolygonDrawArea(103, 104);
            polygon1.SetPoints(new int[] { 11, 12, 13 }, new int[] { 14, 15, 16 });

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(patriarch.Children.Count, 2);

            HSSFPolygon polygon2 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon2.SetPolygonDrawArea(203, 204);
            polygon2.SetPoints(new int[] { 21, 22, 23 }, new int[] { 24, 25, 26 });

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(patriarch.Children.Count, 3);

            polygon = (HSSFPolygon)patriarch.Children[0];
            polygon1 = (HSSFPolygon)patriarch.Children[1];
            polygon2 = (HSSFPolygon)patriarch.Children[2];

            Assert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));
            Assert.AreEqual(polygon.DrawAreaHeight, 101);
            Assert.AreEqual(polygon.DrawAreaWidth, 102);

            Assert.IsTrue(Arrays.Equals(polygon1.XPoints, new int[] { 11, 12, 13 }));
            Assert.IsTrue(Arrays.Equals(polygon1.YPoints, new int[] { 14, 15, 16 }));
            Assert.AreEqual(polygon1.DrawAreaHeight, 104);
            Assert.AreEqual(polygon1.DrawAreaWidth, 103);

            Assert.IsTrue(Arrays.Equals(polygon2.XPoints, new int[] { 21, 22, 23 }));
            Assert.IsTrue(Arrays.Equals(polygon2.YPoints, new int[] { 24, 25, 26 }));
            Assert.AreEqual(polygon2.DrawAreaHeight, 204);
            Assert.AreEqual(polygon2.DrawAreaWidth, 203);
        }
        [Test]
        public void TestExistingFile()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("polygon") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, Drawing.Children.Count);

            HSSFPolygon polygon = (HSSFPolygon)Drawing.Children[0];
            Assert.AreEqual(polygon.DrawAreaHeight, 2466975);
            Assert.AreEqual(polygon.DrawAreaWidth, 3686175);
            Assert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 0, 0, 31479, 16159, 19676, 20502 }));
            Assert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 0, 0, 36, 56, 34, 18 }));
        }
        [Test]
        public void TestPolygonType()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            HSSFPolygon polygon1 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon1.SetPolygonDrawArea(102, 101);
            polygon1.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            EscherSpRecord spRecord = polygon1.GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;

            spRecord.ShapeType = ((short)77/**RANDOM**/);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(patriarch.Children.Count, 2);
            Assert.IsTrue(patriarch.Children[0] is HSSFPolygon);
            Assert.IsTrue(patriarch.Children[1] is HSSFPolygon);
        }
    }

}
