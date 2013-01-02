using System.IO;
using NUnit.Framework;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ooxml.Testcases
{
    /// <summary>
    ///This is a test class for WorksheetDocumentTest and is intended
    ///to contain all WorksheetDocumentTest Unit Tests
    ///</summary>
    [TestFixture]
    public class WorksheetDocumentTest
    {
        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion

        /// <summary>
        ///A test for the Serialization of CT_Worksheet.
        ///</summary>
        [Test]
        public void SerializeEmptyWorksheetDocumentTest()
        {
            CT_Worksheet worksheet = new CT_Worksheet();

            StringWriter stream = new StringWriter();
            WorksheetDocument_Accessor.serializer.Serialize(stream, worksheet, WorksheetDocument_Accessor.namespaces);
            string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
<worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheetData />
</worksheet>";
            Assert.AreEqual(expected, stream.ToString());
        }


        /// <summary>
        ///A test for the Serialization of CT_Worksheet.
        ///</summary>
        [Test]
        public void SerializeWorksheetDocumentTest()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            worksheet.dimension = new CT_SheetDimension();
            worksheet.dimension.@ref = "A1:C1";

            var sheetData = worksheet.AddNewSheetData();
            var row = sheetData.AddNewRow();
            row.r = 1u;
            row.spans = "1:3";
            {
                var c = row.AddNewC();
                c.r = "A1";
                c.t = ST_CellType.s;
                c.v = "0";
            }
            {
                var c = row.AddNewC();
                c.r = "B1";
                c.t = ST_CellType.s;
                c.v = "1";
            }
            {
                var c = row.AddNewC();
                c.r = "C1";
                c.t = ST_CellType.s;
                c.v = "8";
            }
            var hyper = worksheet.AddNewHyperlinks();
            var link = new CT_Hyperlink();
            link.@ref="B1";
            link.id="rId1";
            hyper.hyperlink.Add(link);
            StringWriter stream = new StringWriter();
            WorksheetDocument_Accessor.serializer.Serialize(stream, worksheet, WorksheetDocument_Accessor.namespaces);
            string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
<worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <dimension ref=""A1:C1"" />
  <sheetData>
    <row r=""1"" spans=""1:3"">
      <c r=""A1"" t=""s"">
        <v>0</v>
      </c>
      <c r=""B1"" t=""s"">
        <v>1</v>
      </c>
      <c r=""C1"" t=""s"">
        <v>8</v>
      </c>
    </row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref=""B1"" r:id=""rId1"" />
  </hyperlinks>
</worksheet>";
            Assert.AreEqual(expected, stream.ToString());
        }


        /// <summary>
        ///A test for Deserialization of CT_Worksheet.
        ///</summary>
        [Test]
        public void DeserializeWorksheetDocumentTest()
        {
            // The following is the excerpt of an Excel file.
            string input =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
   <dimension ref=""A1:C10""/>
   <sheetViews>
      <sheetView tabSelected=""1"" view=""pageLayout"" zoomScaleNormal=""100"" workbookViewId=""0"">
         <selection activeCell=""C11"" sqref=""C11""/>
      </sheetView>
   </sheetViews>
   <sheetFormatPr defaultRowHeight=""15""/>
   <sheetData>
      <row r=""1"" spans=""1:3"">
         <c r=""A1"" t=""s"">
            <v>0</v>
         </c>
         <c r=""B1"" t=""s"">
            <v>1</v>
         </c>
         <c r=""C1"" t=""s"">
            <v>8</v>
         </c>
      </row>
      <row r=""2"" spans=""1:3"">
         <c r=""A2"">
            <v>22.3</v>
         </c>
         <c r=""B2"" t=""s"">
            <v>2</v>
         </c>
         <c r=""C2"" t=""s"">
            <v>9</v>
         </c>
      </row>
      <row r=""3"" spans=""1:3"">
         <c r=""A3"">
            <v>24.5</v>
         </c>
         <c r=""B3"" t=""s"">
            <v>3</v>
         </c>
      </row>
      <row r=""4"" spans=""1:3"">
         <c r=""A4"">
            <v>26.7</v>
         </c>
         <c r=""B4"" t=""s"">
            <v>4</v>
         </c>
         <c r=""C4"" s=""1"" t=""s"">
            <v>10</v>
         </c>
      </row>
      <row r=""5"" spans=""1:3"">
         <c r=""A5"">
            <v>41.1</v>
         </c>
         <c r=""B5"" t=""s"">
            <v>5</v>
         </c>
         <c r=""C5"" t=""s"">
            <v>11</v>
         </c>
      </row>
      <row r=""6"" spans=""1:3"">
         <c r=""A6"">
            <v>42.1</v>
         </c>
         <c r=""B6"" t=""s"">
            <v>6</v>
         </c>
      </row>
      <row r=""7"" spans=""1:3"">
         <c r=""A7"">
            <v>42.5</v>
         </c>
         <c r=""B7"" t=""s"">
            <v>7</v>
         </c>
         <c r=""C7"" t=""s"">
            <v>12</v>
         </c>
      </row>
      <row r=""9"" spans=""1:3"">
         <c r=""C9"" t=""s"">
            <v>13</v>
         </c>
      </row>
      <row r=""10"" spans=""1:3"">
         <c r=""C10"" t=""s"">
            <v>14</v>
         </c>
      </row>
   </sheetData>
   <hyperlinks>
      <hyperlink ref=""C4"" r:id=""rId1""/>
   </hyperlinks>
   <pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>
   <pageSetup paperSize=""0"" orientation=""portrait"" horizontalDpi=""0"" verticalDpi=""0"" copies=""0""/>
   <headerFooter>
      <oddHeader>&amp;CThis is the header on sheet 1</oddHeader>
      <oddFooter>&amp;A&amp;RPage &amp;P</oddFooter>
   </headerFooter>
   <legacyDrawing r:id=""rId2""/>
</worksheet>";
            CT_Worksheet result;
            StringReader stream = new StringReader(input);
            result = (CT_Worksheet)WorksheetDocument_Accessor.serializer.Deserialize(stream);

            Assert.AreEqual("A1:C10", result.dimension.@ref);

            Assert.AreEqual(9, result.sheetData.SizeOfRowArray());

      //<row r=""1"" spans=""1:3"">
      //   <c r=""A1"" t=""s"">
      //      <v>0</v>
      //   </c>
      //   <c r=""B1"" t=""s"">
      //      <v>1</v>
      //   </c>
      //   <c r=""C1"" t=""s"">
      //      <v>8</v>
      //   </c>
      //</row>
      //<row r=""2"" spans=""1:3"">
      //   <c r=""A2"">
      //      <v>22.3</v>
      //   </c>
      //   <c r=""B2"" t=""s"">
      //      <v>2</v>
      //   </c>
      //   <c r=""C2"" t=""s"">
      //      <v>9</v>
      //   </c>
      //</row>
            {
                var row = result.sheetData.row[0];
                Assert.AreEqual(1u, row.r);
                Assert.AreEqual("1:3", row.spans);
                Assert.AreEqual("A1", row.c[0].r);
                Assert.AreEqual("0", row.c[0].v);
                Assert.AreEqual("B1", row.c[1].r);
                Assert.AreEqual("1", row.c[1].v);
                Assert.AreEqual("C1", row.c[2].r);
                Assert.AreEqual("8", row.c[2].v);
            }
            {
                var row = result.sheetData.row[1];
                Assert.AreEqual(2u, row.r);
                Assert.AreEqual("1:3", row.spans);
                Assert.AreEqual("A2", row.c[0].r);
                Assert.AreEqual("22.3", row.c[0].v);
                Assert.AreEqual("B2", row.c[1].r);
                Assert.AreEqual("2", row.c[1].v);
                Assert.AreEqual("C2", row.c[2].r);
                Assert.AreEqual("9", row.c[2].v);
            }
        }


        /// <summary>
        ///A test for Deserialization of CT_Worksheet.
        ///</summary>
        [Test]
        public void DeserializeWorksheetDocumentWithPageMarginsTest()
        {
            // The following is the excerpt of an Excel file.
            string input =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
   <dimension ref=""A1""/>
   <sheetViews>
      <sheetView tabSelected=""1"" workbookViewId=""0"">
         <selection activeCell=""A2"" sqref=""A2""/>
      </sheetView>
   </sheetViews>
   <sheetFormatPr defaultRowHeight=""15""/>
   <sheetData>
      <row r=""1"" spans=""1:1"">
         <c r=""A1"" s=""1"">
            <v>1</v>
         </c>
      </row>
   </sheetData>
   <pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>
</worksheet>";
            CT_Worksheet result;
            StringReader stream = new StringReader(input);
            result = (CT_Worksheet)WorksheetDocument_Accessor.serializer.Deserialize(stream);

            Assert.AreEqual("A1", result.dimension.@ref);

            Assert.AreEqual(1, result.sheetData.SizeOfRowArray());

            Assert.IsNotNull(result.sheetFormatPr);
            Assert.AreEqual(15, result.sheetFormatPr.defaultRowHeight);

            {
                var row = result.sheetData.row[0];
                Assert.AreEqual(1u, row.r);
                Assert.AreEqual("1:1", row.spans);

                Assert.AreEqual("A1", row.c[0].r);
                Assert.AreEqual("1", row.c[0].v);
            }

            Assert.IsNotNull(result.pageMargins);
            Assert.AreEqual(0.7, result.pageMargins.left);
            Assert.AreEqual(0.7, result.pageMargins.right);
            Assert.AreEqual(0.75, result.pageMargins.top);
            Assert.AreEqual(0.75, result.pageMargins.bottom);
            Assert.AreEqual(0.3, result.pageMargins.header);
            Assert.AreEqual(0.3, result.pageMargins.footer);
        }

    }
}
