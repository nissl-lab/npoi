using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Model;
using NPOI.DDF;
using TestCases.HSSF.Model;
using NPOI.Util;
using NPOI.HSSF.Record;
using static TestCases.POIFS.Storage.RawDataUtil;

namespace TestCases.HSSF.UserModel
{
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
            polygon.ShapeId = (1024);

            ClassicAssert.AreEqual(polygon.GetEscherContainer().ChildRecords.Count, 4);

            //sp record
            byte[] expected = Decompress("H4sIAAAAAAAAAGNi4PrAwQAELEDMxcAAAAU6ZlwQAAAA");
            byte[] actual = polygon.GetEscherContainer().GetChild(0).Serialize();

            ClassicAssert.AreEqual(expected.Length, actual.Length);
            ClassicAssert.IsTrue(Arrays.Equals(expected, actual));

            expected = Decompress("H4sIAAAAAAAAAGNgEPggxIANAABK4+laGgAAAA==");
            actual = polygon.GetEscherContainer().GetChild(2).Serialize();

            ClassicAssert.AreEqual(expected.Length, actual.Length);
            ClassicAssert.IsTrue(Arrays.Equals(expected, actual));

            expected = Decompress("H4sIAAAAAAAAAGNgEPzAAAQACl6c5QgAAAA=");
            actual = polygon.GetEscherContainer().GetChild(3).Serialize();

            ClassicAssert.AreEqual(expected.Length, actual.Length);
            ClassicAssert.IsTrue(Arrays.Equals(expected, actual));

            ObjRecord obj = polygon.GetObjRecord();

            expected = Decompress("H4sIAAAAAAAAAItlkGIQZRBikGNgYBBMYEADAOAV/ZkeAAAA");
            actual = obj.Serialize();

            ClassicAssert.AreEqual(expected.Length, actual.Length);
            ClassicAssert.IsTrue(Arrays.Equals(expected, actual));

            wb.Close();
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


            EscherArrayProperty verticesProp1 = polygon.GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES) as EscherArrayProperty;

            String expected =
            "<EscherArrayProperty id=\"0x8145\" name=\"geometry.vertices\" blipId=\"false\">" +
            "<Element>[00, 00, 05, 00]</Element>" +
            "<Element>[5A, 00, 05, 00]</Element>" +
            "<Element>[32, 00, 2C, 00]</Element>" +
            "<Element>[5A, 00, 58, 00]</Element>" +
            "<Element>[00, 00, 05, 00]</Element>" +
            "</EscherArrayProperty>";
            String actual = verticesProp1.ToXml("").Replace("\r", "").Replace("\n", "").Replace("\t", "");

            ClassicAssert.AreEqual(verticesProp1.NumberOfElementsInArray, 5);
            ClassicAssert.AreEqual(expected, actual);

            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });
            ClassicAssert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));

            verticesProp1 = polygon.GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES) as EscherArrayProperty;

            expected =
            "<EscherArrayProperty id=\"0x8145\" name=\"geometry.vertices\" blipId=\"false\">" +
            "<Element>[01, 00, 04, 00]</Element>" +
            "<Element>[02, 00, 05, 00]</Element>" +
            "<Element>[03, 00, 06, 00]</Element>" +
            "<Element>[01, 00, 04, 00]</Element>" +
            "</EscherArrayProperty>";
            actual = verticesProp1.ToXml("").Replace("\r", "").Replace("\n", "").Replace("\t", "");

            ClassicAssert.AreEqual(verticesProp1.NumberOfElementsInArray, 4);
            ClassicAssert.AreEqual(expected, actual);

            wb.Close();
        }
        [Test]
        public void TestSetGetProperties()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sh = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            ClassicAssert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));
            ClassicAssert.AreEqual(polygon.DrawAreaHeight, 101);
            ClassicAssert.AreEqual(polygon.DrawAreaWidth, 102);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sh = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            polygon = (HSSFPolygon)patriarch.Children[0];
            ClassicAssert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));
            ClassicAssert.AreEqual(polygon.DrawAreaHeight, 101);
            ClassicAssert.AreEqual(polygon.DrawAreaWidth, 102);

            polygon.SetPolygonDrawArea(1021, 1011);
            polygon.SetPoints(new int[] { 11, 21, 31 }, new int[] { 41, 51, 61 });

            ClassicAssert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 11, 21, 31 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 41, 51, 61 }));
            ClassicAssert.AreEqual(polygon.DrawAreaHeight, 1011);
            ClassicAssert.AreEqual(polygon.DrawAreaWidth, 1021);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();
            sh = wb3.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            polygon = (HSSFPolygon)patriarch.Children[0];

            ClassicAssert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 11, 21, 31 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 41, 51, 61 }));
            ClassicAssert.AreEqual(polygon.DrawAreaHeight, 1011);
            ClassicAssert.AreEqual(polygon.DrawAreaWidth, 1021);

            wb3.Close();
        }
        [Test]
        public void TestAddToExistingFile()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sh = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            HSSFPolygon polygon1 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon1.SetPolygonDrawArea(103, 104);
            polygon1.SetPoints(new int[] { 11, 12, 13 }, new int[] { 14, 15, 16 });

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sh = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(patriarch.Children.Count, 2);

            HSSFPolygon polygon2 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon2.SetPolygonDrawArea(203, 204);
            polygon2.SetPoints(new int[] { 21, 22, 23 }, new int[] { 24, 25, 26 });

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();
            sh = wb3.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(patriarch.Children.Count, 3);

            polygon = (HSSFPolygon)patriarch.Children[0];
            polygon1 = (HSSFPolygon)patriarch.Children[1];
            polygon2 = (HSSFPolygon)patriarch.Children[2];

            ClassicAssert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 1, 2, 3 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 4, 5, 6 }));
            ClassicAssert.AreEqual(polygon.DrawAreaHeight, 101);
            ClassicAssert.AreEqual(polygon.DrawAreaWidth, 102);

            ClassicAssert.IsTrue(Arrays.Equals(polygon1.XPoints, new int[] { 11, 12, 13 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon1.YPoints, new int[] { 14, 15, 16 }));
            ClassicAssert.AreEqual(polygon1.DrawAreaHeight, 104);
            ClassicAssert.AreEqual(polygon1.DrawAreaWidth, 103);

            ClassicAssert.IsTrue(Arrays.Equals(polygon2.XPoints, new int[] { 21, 22, 23 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon2.YPoints, new int[] { 24, 25, 26 }));
            ClassicAssert.AreEqual(polygon2.DrawAreaHeight, 204);
            ClassicAssert.AreEqual(polygon2.DrawAreaWidth, 203);

            wb3.Close();
        }
        [Test]
        public void TestExistingFile()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("polygon") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            ClassicAssert.AreEqual(1, Drawing.Children.Count);

            HSSFPolygon polygon = (HSSFPolygon)Drawing.Children[0];
            ClassicAssert.AreEqual(polygon.DrawAreaHeight, 2466975);
            ClassicAssert.AreEqual(polygon.DrawAreaWidth, 3686175);
            ClassicAssert.IsTrue(Arrays.Equals(polygon.XPoints, new int[] { 0, 0, 31479, 16159, 19676, 20502 }));
            ClassicAssert.IsTrue(Arrays.Equals(polygon.YPoints, new int[] { 0, 0, 36, 56, 34, 18 }));

            wb.Close();
        }
        [Test]
        public void TestPolygonType()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sh = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPolygonDrawArea(102, 101);
            polygon.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sh = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            HSSFPolygon polygon1 = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon1.SetPolygonDrawArea(102, 101);
            polygon1.SetPoints(new int[] { 1, 2, 3 }, new int[] { 4, 5, 6 });

            EscherSpRecord spRecord = polygon1.GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;

            spRecord.ShapeType = ((short)77/**RANDOM**/);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();
            sh = wb3.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(patriarch.Children.Count, 2);
            ClassicAssert.IsTrue(patriarch.Children[0] is HSSFPolygon);
            ClassicAssert.IsTrue(patriarch.Children[1] is HSSFPolygon);

            wb3.Close();
        }
    }

}
