using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCases.HSSF.UserModel
{

    /**
     * @author Evgeniy Berlog
     * @date 28.06.12
     */
    public class TestPolygon
    {

        public void TestResultEqualsToAbstractShape()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(100, 100);
            polygon.SetPoints(new int[] { 0, 90, 50 }, new int[] { 5, 5, 44 });
            PolygonShape polygonShape = HSSFTestModelHelper.CreatePolygonShape(1024, polygon);
            polygon.SetShapeId(1024);

            Assert.AreEqual(polygon.GetEscherContainer().GetChildRecords().Count, 4);
            Assert.AreEqual(polygonShape.GetSpContainer().GetChildRecords().Count, 4);

            //sp record
            byte[] expected = polygonShape.GetSpContainer().GetChild(0).Serialize();
            byte[] actual = polygon.GetEscherContainer().GetChild(0).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = polygonShape.GetSpContainer().GetChild(2).Serialize();
            actual = polygon.GetEscherContainer().GetChild(2).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = polygonShape.GetSpContainer().GetChild(3).Serialize();
            actual = polygon.GetEscherContainer().GetChild(3).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            ObjRecord obj = polygon.GetObjRecord();
            ObjRecord objShape = polygonShape.GetObjRecord();

            expected = obj.Serialize();
            actual = objShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));
        }

        public void TestPolygonPoints()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(100, 100);
            polygon.SetPoints(new int[] { 0, 90, 50, 90 }, new int[] { 5, 5, 44, 88 });

            PolygonShape polygonShape = HSSFTestModelHelper.CreatePolygonShape(0, polygon);

            EscherArrayProperty verticesProp1 = polygon.GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES);
            EscherArrayProperty verticesProp2 = ((EscherOptRecord)polygonShape.GetSpContainer().GetChildById(EscherOptRecord.RECORD_ID))
                    .Lookup(EscherProperties.GEOMETRY__VERTICES);

            Assert.AreEqual(verticesProp1.GetNumberOfElementsInArray(), verticesProp2.GetNumberOfElementsInArray());
            Assert.AreEqual(verticesProp1.ToXml(""), verticesProp2.ToXml(""));

            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });
            Assert.IsTrue(Arrays.Equals(polygon.GetXPoints(), new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.GetYPoints(), new int[] { 4, 5, 6 }));

            polygonShape = HSSFTestModelHelper.CreatePolygonShape(0, polygon);
            verticesProp1 = polygon.GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES);
            verticesProp2 = ((EscherOptRecord)polygonShape.GetSpContainer().GetChildById(EscherOptRecord.RECORD_ID))
                    .Lookup(EscherProperties.GEOMETRY__VERTICES);

            Assert.AreEqual(verticesProp1.GetNumberOfElementsInArray(), verticesProp2.GetNumberOfElementsInArray());
            Assert.AreEqual(verticesProp1.ToXml(""), verticesProp2.ToXml(""));
        }

        public void TestSetGetProperties()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            Assert.IsTrue(Arrays.Equals(polygon.GetXPoints(), new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.GetYPoints(), new int[] { 4, 5, 6 }));
            Assert.AreEqual(polygon.GetDrawAreaHeight(), 101);
            Assert.AreEqual(polygon.GetDrawAreaWidth(), 102);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            polygon = (HSSFPolygon)patriarch.Children().Get(0);
            Assert.IsTrue(Arrays.Equals(polygon.GetXPoints(), new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.GetYPoints(), new int[] { 4, 5, 6 }));
            Assert.AreEqual(polygon.GetDrawAreaHeight(), 101);
            Assert.AreEqual(polygon.GetDrawAreaWidth(), 102);

            polygon.SetPolygonDrawArea(1021, 1011);
            polygon.SetPoints(new int[] { 11, 21, 31 }, new int[] { 41, 51, 61 });

            Assert.IsTrue(Arrays.Equals(polygon.GetXPoints(), new int[] { 11, 21, 31 }));
            Assert.IsTrue(Arrays.Equals(polygon.GetYPoints(), new int[] { 41, 51, 61 }));
            Assert.AreEqual(polygon.GetDrawAreaHeight(), 1011);
            Assert.AreEqual(polygon.GetDrawAreaWidth(), 1021);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            polygon = (HSSFPolygon)patriarch.Children().Get(0);

            Assert.IsTrue(Arrays.Equals(polygon.GetXPoints(), new int[] { 11, 21, 31 }));
            Assert.IsTrue(Arrays.Equals(polygon.GetYPoints(), new int[] { 41, 51, 61 }));
            Assert.AreEqual(polygon.GetDrawAreaHeight(), 1011);
            Assert.AreEqual(polygon.GetDrawAreaWidth(), 1021);
        }

        public void TestAddToExistingFile()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            HSSFPolygon polygon1 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon1.SetPolygonDrawArea(103, 104);
            polygon1.SetPoints(new int[] { 11, 12, 13 }, new int[] { 14, 15, 16 });

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            Assert.AreEqual(patriarch.Children().Count, 2);

            HSSFPolygon polygon2 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon2.SetPolygonDrawArea(203, 204);
            polygon2.SetPoints(new int[] { 21, 22, 23 }, new int[] { 24, 25, 26 });

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            Assert.AreEqual(patriarch.Children().Count, 3);

            polygon = (HSSFPolygon)patriarch.Children().Get(0);
            polygon1 = (HSSFPolygon)patriarch.Children().Get(1);
            polygon2 = (HSSFPolygon)patriarch.Children().Get(2);

            Assert.IsTrue(Arrays.Equals(polygon.GetXPoints(), new int[] { 1, 2, 3 }));
            Assert.IsTrue(Arrays.Equals(polygon.GetYPoints(), new int[] { 4, 5, 6 }));
            Assert.AreEqual(polygon.GetDrawAreaHeight(), 101);
            Assert.AreEqual(polygon.GetDrawAreaWidth(), 102);

            Assert.IsTrue(Arrays.Equals(polygon1.GetXPoints(), new int[] { 11, 12, 13 }));
            Assert.IsTrue(Arrays.Equals(polygon1.GetYPoints(), new int[] { 14, 15, 16 }));
            Assert.AreEqual(polygon1.GetDrawAreaHeight(), 104);
            Assert.AreEqual(polygon1.GetDrawAreaWidth(), 103);

            Assert.IsTrue(Arrays.Equals(polygon2.GetXPoints(), new int[] { 21, 22, 23 }));
            Assert.IsTrue(Arrays.Equals(polygon2.GetYPoints(), new int[] { 24, 25, 26 }));
            Assert.AreEqual(polygon2.GetDrawAreaHeight(), 204);
            Assert.AreEqual(polygon2.GetDrawAreaWidth(), 203);
        }

        public void TestExistingFile()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("polygon");
            HSSFPatriarch Drawing = sheet.DrawingPatriarch();
            Assert.AreEqual(1, Drawing.Children().Count);

            HSSFPolygon polygon = (HSSFPolygon)Drawing.Children().Get(0);
            Assert.AreEqual(polygon.GetDrawAreaHeight(), 2466975);
            Assert.AreEqual(polygon.GetDrawAreaWidth(), 3686175);
            Assert.IsTrue(Arrays.Equals(polygon.GetXPoints(), new int[] { 0, 0, 31479, 16159, 19676, 20502 }));
            Assert.IsTrue(Arrays.Equals(polygon.GetYPoints(), new int[] { 0, 0, 36, 56, 34, 18 }));
        }

        public void TestPolygonType()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            HSSFPolygon polygon1 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon1.SetPolygonDrawArea(102, 101);
            polygon1.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            EscherSpRecord spRecord = polygon1.GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);

            spRecord.SetShapeType((short)77/**RANDOM**/);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            Assert.AreEqual(patriarch.Children().Count, 2);
            Assert.IsTrue(patriarch.Children().Get(0) is HSSFPolygon);
            Assert.IsTrue(patriarch.Children().Get(1) is HSSFPolygon);
        }
    }

}
