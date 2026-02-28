using System.IO;
using NUnit.Framework;
using NPOI.OpenXmlFormats.Dml;

namespace ooxml.Testcases
{
    /// <summary>
    ///This is a test class for VmlDrawingDocumentTest and is intended
    ///to contain all VmlDrawingDocumentTest Unit Tests
    ///</summary>
    [TestFixture]
    public class VmlDrawingDocumentTest
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


//        /// <summary>
//        ///A test for the Serialization of CT_Drawing.
//        ///</summary>
//        [Test]
//        public void SerializeVmlDrawingDocumentTest()
//        {
//            var drawing = new NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing();

//            using (StringWriter stream = new StringWriter())
//            {
//                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing_Accessor.serializer.Serialize(stream, drawing, NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing_Accessor.namespaces);
//                string expected = @"<xml xmlns:v=""urn:schemas-microsoft-com:vml""
// xmlns:o=""urn:schemas-microsoft-com:office:office""
// xmlns:x=""urn:schemas-microsoft-com:office:excel"">
//<xdr:wsDr xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" />";

//                //</xdr:wsDr>";
//                Assert.AreEqual(expected, stream.ToString());
//            }
//        }

        /// <summary>
        ///A test for Deserialize
        ///</summary>
        [Test]
        public void DeserializeVmlDrawingDocumentTest()
        {
            // The following is the excerpt of an Excel file.
            string input =
@"<xml xmlns:v=""urn:schemas-microsoft-com:vml""
 xmlns:o=""urn:schemas-microsoft-com:office:office""
 xmlns:x=""urn:schemas-microsoft-com:office:excel"">
 <o:shapelayout v:ext=""edit"">
  <o:idmap v:ext=""edit"" data=""3""/>
 </o:shapelayout><v:shapetype id=""_x0000_t201"" coordsize=""21600,21600"" o:spt=""201""
  path=""m,l,21600r21600,l21600,xe"">
  <v:stroke joinstyle=""miter""/>
  <v:path shadowok=""f"" o:extrusionok=""f"" strokeok=""f"" fillok=""f"" o:connecttype=""rect""/>
  <o:lock v:ext=""edit"" shapetype=""t""/>
 </v:shapetype><v:shape id=""CreatePNList"" o:spid=""_x0000_s3093"" type=""#_x0000_t201""
  style='position:absolute;margin-left:894pt;margin-top:56.25pt;width:96.75pt;
  height:31.5pt;z-index:1' fillcolor=""white [9]"" stroked=""f"" strokecolor=""windowText [64]""
  strokeweight=""3e-5mm"" o:insetmode=""auto"">
  <v:fill color2=""white [9]""/>
  <v:imagedata o:relid=""rId1"" o:title=""""/>
  <v:shadow color=""black"" obscured=""t""/>
  <x:ClientData ObjectType=""Pict"">
   <x:SizeWithCells/>
   <x:Anchor>
    19, 0, 1, 44, 20, 2, 3, 6</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:CF>Pict</x:CF>
   <x:AutoPict/>
  </x:ClientData>
 </v:shape><v:shapetype id=""_x0000_t202"" coordsize=""21600,21600"" o:spt=""202""
  path=""m,l,21600r21600,l21600,xe"">
  <v:stroke joinstyle=""miter""/>
  <v:path gradientshapeok=""t"" o:connecttype=""rect""/>
 </v:shapetype><v:shape id=""_x0000_s3147"" type=""#_x0000_t202"" style='position:absolute;
  margin-left:93.75pt;margin-top:715.5pt;width:95.25pt;height:28.5pt;z-index:2;
  visibility:visible;mso-wrap-style:tight' fillcolor=""#ffffe1"" o:insetmode=""auto"">
  <v:fill color2=""#ffffe1""/>
  <v:shadow on=""t"" color=""black"" obscured=""t""/>
  <v:path o:connecttype=""none""/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType=""Note"">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    2, 0, 51, 6, 3, 117, 53, 10</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>51</x:Row>
   <x:Column>1</x:Column>
   <x:Visible/>
  </x:ClientData>
 </v:shape><v:shape id=""_x0000_s3148"" type=""#_x0000_t202"" style='position:absolute;
  margin-left:105.75pt;margin-top:741.75pt;width:94.5pt;height:28.5pt;
  z-index:3;visibility:hidden;mso-wrap-style:tight' fillcolor=""#ffffe1""
  o:insetmode=""auto"">
  <v:fill color2=""#ffffe1""/>
  <v:shadow on=""t"" color=""black"" obscured=""t""/>
  <v:path o:connecttype=""none""/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType=""Note"">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    3, 6, 53, 7, 4, 1, 55, 11</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>52</x:Row>
   <x:Column>1</x:Column>
  </x:ClientData>
 </v:shape><v:shape id=""_x0000_s3150"" type=""#_x0000_t202"" style='position:absolute;
  margin-left:1221pt;margin-top:815.25pt;width:260.25pt;height:22.5pt;
  z-index:4;visibility:visible;mso-wrap-style:tight' fillcolor=""#ffffe1""
  o:insetmode=""auto"">
  <v:fill color2=""#ffffe1""/>
  <v:shadow on=""t"" color=""black"" obscured=""t""/>
  <v:path o:connecttype=""none""/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType=""Note"">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    21, 159, 59, 3, 21, 506, 60, 16</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>57</x:Row>
   <x:Column>21</x:Column>
   <x:Visible/>
  </x:ClientData>
 </v:shape><v:shape id=""_x0000_s3152"" type=""#_x0000_t202"" style='position:absolute;
  margin-left:1603.5pt;margin-top:145.5pt;width:290.25pt;height:35.25pt;
  z-index:5;visibility:visible;mso-wrap-style:tight' fillcolor=""#ffffe1""
  o:insetmode=""auto"">
  <v:fill color2=""#ffffe1""/>
  <v:shadow on=""t"" color=""black"" obscured=""t""/>
  <v:path o:connecttype=""none""/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType=""Note"">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    22, 139, 6, 11, 24, 226, 9, 7</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>7</x:Row>
   <x:Column>22</x:Column>
   <x:Visible/>
  </x:ClientData>
 </v:shape></xml>";
            //NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing result;
            ////{
            ////    StringReader stream = new StringReader(input);
            ////    result = (CT_Drawing)CommentsDocument_Accessor.serializer.Deserialize(stream); // instantiate source code to enable debugging the serialization code
            ////}
            //{
            //    using (StringReader stream = new StringReader(input))
            //    {
            //        result = (NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing)NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing_Accessor.serializer.Deserialize(stream);
            //    }
            //}
            //Assert.IsNotNull(result.TwoCellAnchors);
            //Assert.AreEqual(1, result.TwoCellAnchors.Count);
            //Assert.AreEqual(1, result.SizeOfTwoCellAnchorArray());
            //var anchor = result.TwoCellAnchors[0];
            //Assert.AreEqual(0, anchor.from.col);
        }



        /// <summary>
        ///A test for Deserialize
        ///</summary>
        [Test]
        public void DeserializeVmlDrawingDocumentWithPicturesTest()
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
