using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using System;
using System.IO;

namespace ooxml.Testcases
{
    
    
    /// <summary>
    ///This is a test class for WorkbookDocumentTest and is intended
    ///to contain all WorkbookDocumentTest Unit Tests
    ///</summary>
    [TestFixture]
    public class WorkbookDocumentTest
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
            CT_Workbook worksheet = new CT_Workbook();

            StringWriter stream = new StringWriter();
            WorkbookDocument_Accessor.serializer.Serialize(stream, worksheet, WorkbookDocument_Accessor.namespaces);
            string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
<workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheets />
</workbook>";
            Assert.AreEqual(expected, stream.ToString());
        }


        /// <summary>
        ///A test for the Serialization of CT_Worksheet.
        ///</summary>
        [Test]
        public void SerializeWorksheetDocumentTest()
        {
            CT_Workbook worksheet = new CT_Workbook();
            CT_Sheet sheet1 = new CT_Sheet();
            sheet1.name = "Sheet1";
            sheet1.sheetId = 1u;
            sheet1.id = "rId1";
            worksheet.sheets.sheet.Add(sheet1);

            var bks = worksheet.AddNewBookViews();
            var bk = bks.AddNewWorkbookView();
            bk.xWindow = 360;
            bk.xWindowSpecified = true;
            bk.yWindow = 60;
            bk.yWindowSpecified = true;
            bk.windowWidth = 11295;
            bk.windowWidthSpecified = true;
            bk.windowHeight = 5580;
            bk.windowHeightSpecified = true;



            StringWriter stream = new StringWriter();
            WorkbookDocument_Accessor.serializer.Serialize(stream, worksheet, WorkbookDocument_Accessor.namespaces);
            string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
<workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <bookViews>
    <workbookView xWindow=""360"" yWindow=""60"" windowWidth=""11295"" windowHeight=""5580"" />
  </bookViews>
  <sheets>
    <sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1"" />
  </sheets>
</workbook>";
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
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
   <fileVersion appName=""xl"" lastEdited=""4"" lowestEdited=""4"" rupBuild=""4505""/>
   <workbookPr defaultThemeVersion=""124226""/>
   <bookViews>
      <workbookView xWindow=""360"" yWindow=""60"" windowWidth=""11295"" windowHeight=""5580""/>
   </bookViews>
   <sheets>
      <sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/>
      <sheet name=""Sheet2"" sheetId=""2"" r:id=""rId2""/>
      <sheet name=""Sheet3"" sheetId=""3"" r:id=""rId3""/>
   </sheets>
   <definedNames>
      <definedName name=""AllANumbers"" comment=""All the numbers in A"">Sheet1!$A$2:$A$7</definedName>
      <definedName name=""AllBStrings"" comment=""All the strings in B"">Sheet1!$B$2:$B$7</definedName>
   </definedNames>
   <calcPr calcId=""125725""/>
</workbook>";
            CT_Workbook result;
            StringReader stream = new StringReader(input);
            result = (CT_Workbook)WorkbookDocument_Accessor.serializer.Deserialize(stream);

            Assert.AreEqual("Sheet1", result.sheets.sheet[0].name);
            Assert.AreEqual(1u, result.sheets.sheet[0].sheetId);
            Assert.AreEqual("rId1", result.sheets.sheet[0].id);

            Assert.AreEqual("Sheet2", result.sheets.sheet[1].name);
            Assert.AreEqual(2u, result.sheets.sheet[1].sheetId);
            Assert.AreEqual("rId2", result.sheets.sheet[1].id);

            Assert.AreEqual("Sheet3", result.sheets.sheet[2].name);
            Assert.AreEqual(3u, result.sheets.sheet[2].sheetId);
            Assert.AreEqual("rId3", result.sheets.sheet[2].id);
        }


    }
}
