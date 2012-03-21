using NPOI.OpenXmlFormats.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace ooxml.Testcases
{


    /// <summary>
    ///This is a test class for CommentsDocumentTest and is intended
    ///to contain all CommentsDocumentTest Unit Tests
    ///</summary>
    [TestClass()]
    public class CommentsDocumentTest
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
        ///A test for Serialize
        ///</summary>
        [TestMethod()]
        public void SerializeCommentsDocumentTest()
        {
            CT_Comments comments = new CT_Comments();
            comments.AddNewAuthors().AddAuthor("Christian");
            comments.authors.AddAuthor("Tony");
            comments.AddNewCommentList();
            var text = new CT_Rst();
            text.t = "First Comment";
            comments.commentList.AddNewComment().text = text;           
            StringWriter stream = new StringWriter();
            CommentsDocument_Accessor.serializer.Serialize(stream, comments);
            string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
<comments xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <authors>
    <author>
      <author>Christian</author>
      <author>Tony</author>
    </author>
  </authors>
  <commentList>
    <comment />
  </commentList>
  <extLst>
    <ext />
  </extLst>
</comments>";
            Assert.AreEqual(expected, stream.ToString());
        }


        /// <summary>
        ///A test for Deserialize
        ///</summary>
        [Ignore] // TODO fix
        [TestMethod()]
        public void DeserializeCommentsDocumentTest()
        {
            // The following is the excerpt of an Excel file.
            string input =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<comments xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
   <authors>
      <author>Autor</author>
   </authors>
   <commentList>
      <comment ref=""A3"" authorId=""0"">
         <text>
            <r>
               <rPr>
                  <b/>
                  <sz val=""9""/>
                  <color indexed=""81""/>
                  <rFont val=""Tahoma""/>
                  <charset val=""1""/>
               </rPr>
               <t>Autor:</t>
            </r>
            <r>
               <rPr>
                  <sz val=""9""/>
                  <color indexed=""81""/>
                  <rFont val=""Tahoma""/>
                  <charset val=""1""/>
               </rPr>
               <t xml:space=""preserve"">
First Column on the first sheet; This is just an increasing column</t>
            </r>
         </text>
      </comment>
      <comment ref=""B3"" authorId=""0"">
         <text>
            <r>
               <rPr>
                  <b/>
                  <sz val=""9""/>
                  <color indexed=""81""/>
                  <rFont val=""Tahoma""/>
                  <charset val=""1""/>
               </rPr>
               <t>Autor:</t>
            </r>
            <r>
               <rPr>
                  <sz val=""9""/>
                  <color indexed=""81""/>
                  <rFont val=""Tahoma""/>
                  <charset val=""1""/>
               </rPr>
               <t xml:space=""preserve"">
This column multiplies the first column value with 2.</t>
            </r>
         </text>
      </comment>
      <comment ref=""C3"" authorId=""0"">
         <text>
            <r>
               <rPr>
                  <b/>
                  <sz val=""9""/>
                  <color indexed=""81""/>
                  <rFont val=""Tahoma""/>
                  <charset val=""1""/>
               </rPr>
               <t>Autor:</t>
            </r>
            <r>
               <rPr>
                  <sz val=""9""/>
                  <color indexed=""81""/>
                  <rFont val=""Tahoma""/>
                  <charset val=""1""/>
               </rPr>
               <t xml:space=""preserve"">
This column adds the first and second one.</t>
            </r>
         </text>
      </comment>
   </commentList>
</comments>";
            CT_Comments result;
            {
                StringReader stream = new StringReader(input);
                result = (CT_Comments)CommentsDocument_Accessor.serializer.Deserialize(stream); // instiate of debugging
            }
            {
                StringReader stream = new StringReader(input);
                result = (CT_Comments)CommentsDocument_Accessor.serializer.Deserialize(stream);
            }
            Assert.AreEqual(3, result.commentList.SizeOfCommentArray());
            Assert.AreEqual(1, result.authors.SizeOfAuthorArray());
            Assert.AreEqual("First Column on the first sheet; This is just an increasing column", result.commentList.GetCommentArray(0).text.t);
        }



    }
}
