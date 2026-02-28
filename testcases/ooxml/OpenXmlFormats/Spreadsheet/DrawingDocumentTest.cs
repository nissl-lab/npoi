using System.IO;
using NUnit.Framework;
using NPOI.OpenXmlFormats.Dml;

namespace ooxml.Testcases
{
    /// <summary>
    ///This is a test class for DrawingDocumentTest and is intended
    ///to contain all DrawingDocumentTest Unit Tests
    ///</summary>
    [TestFixture]
    public class DrawingDocumentTest
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
        ///A test for the Serialization of CT_Drawing.
        ///</summary>
        [Test]
        public void SerializeDrawingDocumentTest()
        {
            var drawing = new NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing();
            
            using (StringWriter stream = new StringWriter())
            {
                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing_Accessor.serializer.Serialize(stream, drawing, NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing_Accessor.namespaces);
                string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
<xdr:wsDr xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" />";

//</xdr:wsDr>";
                Assert.AreEqual(expected, stream.ToString());
            }
        }

        /// <summary>
        ///A test for Deserialize
        ///</summary>
        [Test]
        public void DeserializeDrawingDocumentTest()
        {
            // The following is the excerpt of an Excel file.
            string input =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
   <xdr:twoCellAnchor>
      <xdr:from>
         <xdr:col>0</xdr:col>
         <xdr:colOff>38100</xdr:colOff>
         <xdr:row>17</xdr:row>
         <xdr:rowOff>28575</xdr:rowOff>
      </xdr:from>
      <xdr:to>
         <xdr:col>5</xdr:col>
         <xdr:colOff>600075</xdr:colOff>
         <xdr:row>44</xdr:row>
         <xdr:rowOff>0</xdr:rowOff>
      </xdr:to>
      <xdr:graphicFrame macro="""">
         <xdr:nvGraphicFramePr>
            <xdr:cNvPr id=""1027"" name=""Chart 3""/>
            <xdr:cNvGraphicFramePr>
               <a:graphicFrameLocks/>
            </xdr:cNvGraphicFramePr>
         </xdr:nvGraphicFramePr>
         <xdr:xfrm>
            <a:off x=""0"" y=""0""/>
            <a:ext cx=""0"" cy=""0""/>
         </xdr:xfrm>
         <a:graphic>
            <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
               <c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" r:id=""rId1""/>
            </a:graphicData>
         </a:graphic>
      </xdr:graphicFrame>
      <xdr:clientData/>
   </xdr:twoCellAnchor>
</xdr:wsDr>";
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing result;
            //{
            //    StringReader stream = new StringReader(input);
            //    result = (CT_Drawing)CommentsDocument_Accessor.serializer.Deserialize(stream); // instantiate source code to enable debugging the serialization code
            //}
            {
                using (StringReader stream = new StringReader(input))
                {
                    result = (NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing)NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing_Accessor.serializer.Deserialize(stream);
                }
            }
            Assert.IsNotNull(result.TwoCellAnchors);
            Assert.AreEqual(1, result.TwoCellAnchors.Count);
            Assert.AreEqual(1, result.SizeOfTwoCellAnchorArray());
            var anchor = result.TwoCellAnchors[0];
            Assert.AreEqual(0, anchor.from.col);
        }



        /// <summary>
        ///A test for Deserialize
        ///</summary>
        [Test]
        public void DeserializeDrawingDocumentWithPicturesTest()
        {
            // The following is the excerpt of an Excel file.
            string input =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<xdr:wsDr xmlns:fo=""urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"" xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">
   <xdr:twoCellAnchor>
      <xdr:from>
         <xdr:col>2</xdr:col>
         <xdr:colOff>0</xdr:colOff>
         <xdr:row>0</xdr:row>
         <xdr:rowOff>0</xdr:rowOff>
      </xdr:from>
      <xdr:to>
         <xdr:col>20</xdr:col>
         <xdr:colOff>544708</xdr:colOff>
         <xdr:row>42</xdr:row>
         <xdr:rowOff>161666</xdr:rowOff>
      </xdr:to>
      <xdr:pic>
         <xdr:nvPicPr>
            <xdr:cNvPr id=""1"" name=""Graphics 1"" descr=""Graphics 1"" />
            <xdr:cNvPicPr>
               <a:picLocks noChangeAspect=""0"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" />
            </xdr:cNvPicPr>
         </xdr:nvPicPr>
         <xdr:blipFill>
            <a:blip r:embed=""Monozipmx003Ammx002Fmmx002Fmlocalhostmx002Fmcontent.xmlElementm0m2m3m0m0m3m1m0m"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" />
            <a:stretch xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
               <a:fillRect />
            </a:stretch>
         </xdr:blipFill>
         <xdr:spPr>
            <a:xfrm xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" />
            <a:prstGeom prst=""rect"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
               <a:avLst />
            </a:prstGeom>
            <a:noFill xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" />
            <a:ln xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
               <a:noFill />
            </a:ln>
         </xdr:spPr>
      </xdr:pic>
      <xdr:clientData />
   </xdr:twoCellAnchor>
</xdr:wsDr>";
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing result;
            //{
            //    StringReader stream = new StringReader(input);
            //    result = (CT_Drawing)CommentsDocument_Accessor.serializer.Deserialize(stream); // instantiate source code to enable debugging the serialization code
            //}
            {
                using (StringReader stream = new StringReader(input))
                {
                    result = (NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing)NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing_Accessor.serializer.Deserialize(stream);
                }
            }
            Assert.IsNotNull(result.TwoCellAnchors);
            Assert.AreEqual(1, result.TwoCellAnchors.Count);
            Assert.AreEqual(1, result.SizeOfTwoCellAnchorArray());
            var anchor = result.TwoCellAnchors[0];
            Assert.AreEqual(2, anchor.from.col);
        }


    }
}
