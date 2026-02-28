/* ====================================================================
  Licensed to the Apache Software Foundation (ASF) under one or more
  contributor license agreements.  See the NOTICE file distributed with
  this work for Additional information regarding copyright ownership.
  The ASF licenses this file to You under the Apache License, Version 2.0
  (the "License"); you may not use this file except in compliance with
  the License.  You may obtain a copy of the License at

      http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
==================================================================== */

namespace TestCases
{

    using NPOI;
    using NPOI.OpenXmlFormats;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using TestCases.XWPF;

    /**
     * Test Setting extended and custom OOXML properties
     */
    [TestFixture]
    public class TestPOIXMLProperties
    {
        private XWPFDocument sampleDoc;
        private XWPFDocument sampleNoThumb;
        private POIXMLProperties _props;
        private CoreProperties _coreProperties;

        [SetUp]
        public void SetUp()
        {
            sampleDoc = XWPFTestDataSamples.OpenSampleDocument("documentProperties.docx");
            sampleNoThumb = XWPFTestDataSamples.OpenSampleDocument("SampleDoc.docx");
            ClassicAssert.IsNotNull(sampleDoc);
            ClassicAssert.IsNotNull(sampleNoThumb);
            _props = sampleDoc.GetProperties();
            _coreProperties = _props.CoreProperties;
            ClassicAssert.IsNotNull(_props);
        }


        [TearDown]
        public void closeResources()
        {
            sampleDoc.Close();
            sampleNoThumb.Close();
        }

        [Test]
        public void TestWorkbookExtendedProperties()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            POIXMLProperties props = workbook.GetProperties();
            ClassicAssert.IsNotNull(props);

            ExtendedProperties properties =
                    props.ExtendedProperties;

            CT_ExtendedProperties
                    ctProps = properties.GetUnderlyingProperties();


            String appVersion = "3.5 beta";
            String application = "POI";

            ctProps.Application = (application);
            ctProps.AppVersion = (appVersion);

            XSSFWorkbook newWorkbook =
                    (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            workbook.Close();
            ClassicAssert.IsTrue(workbook != newWorkbook);


            POIXMLProperties newProps = newWorkbook.GetProperties();
            ClassicAssert.IsNotNull(newProps);
            ExtendedProperties newProperties =
                    newProps.ExtendedProperties;

            ClassicAssert.AreEqual(application, newProperties.Application);
            ClassicAssert.AreEqual(appVersion, newProperties.AppVersion);
        

            CT_ExtendedProperties
                    newCtProps = newProperties.GetUnderlyingProperties();

            ClassicAssert.AreEqual(application, newCtProps.Application);
            ClassicAssert.AreEqual(appVersion, newCtProps.AppVersion);

            newWorkbook.Close();
        }


        /**
         * Test usermodel API for Setting custom properties
         */
        [Test]
        public void TestCustomProperties()
        {
            POIXMLDocument wb1 = new XSSFWorkbook();

            CustomProperties customProps = wb1.GetProperties().CustomProperties;
            customProps.AddProperty("test-1", "string val");
            customProps.AddProperty("test-2", 1974);
            customProps.AddProperty("test-3", 36.6);
            //Adding a duplicate
            try
            {
                customProps.AddProperty("test-3", 36.6);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.AreEqual("A property with this name already exists in the custom properties", e.Message);
            }
            customProps.AddProperty("test-4", true);

            POIXMLDocument wb2 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack((XSSFWorkbook)wb1);
            wb1.Close();

            CT_CustomProperties ctProps =
                    wb2.GetProperties().CustomProperties.GetUnderlyingProperties();
            ClassicAssert.AreEqual(6, ctProps.sizeOfPropertyArray());
            CT_Property p;

            p = ctProps.GetPropertyArray(0);
            ClassicAssert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            ClassicAssert.AreEqual("test-1", p.name);
            ClassicAssert.AreEqual("string val", p.Item.ToString());
            ClassicAssert.AreEqual(2, p.pid);

            p = ctProps.GetPropertyArray(1);
            ClassicAssert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            ClassicAssert.AreEqual("test-2", p.name);
            ClassicAssert.AreEqual(1974, p.Item);
            ClassicAssert.AreEqual(3, p.pid);

            p = ctProps.GetPropertyArray(2);
            ClassicAssert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            ClassicAssert.AreEqual("test-3", p.name);
            ClassicAssert.AreEqual(36.6, p.Item);
            ClassicAssert.AreEqual(4, p.pid);

            p = ctProps.GetPropertyArray(3);
            ClassicAssert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            ClassicAssert.AreEqual("test-4", p.name);
            ClassicAssert.AreEqual(true, p.Item);
            ClassicAssert.AreEqual(5, p.pid);

            p = ctProps.GetPropertyArray(4);
            ClassicAssert.AreEqual("Generator", p.name);
            ClassicAssert.AreEqual("NPOI", p.Item);
            ClassicAssert.AreEqual(6, p.pid);

            //p = ctProps.GetPropertyArray(5);
            //ClassicAssert.AreEqual("Generator Version", p.name);
            //ClassicAssert.AreEqual("2.0.9", p.Item);
            //ClassicAssert.AreEqual(7, p.pid);

            wb2.Close();
        }
        [Ignore("test")]
        public void TestDocumentProperties()
        {
            String category = _coreProperties.Category;
            ClassicAssert.AreEqual("test", category);
            String contentStatus = "Draft";
            _coreProperties.ContentStatus = contentStatus;
            ClassicAssert.AreEqual("Draft", contentStatus);
            DateTime? Created = _coreProperties.Created;
            // the original file Contains a following value: 2009-07-20T13:12:00Z
            ClassicAssert.IsTrue(DateTimeEqualToUTCString(Created, "2009-07-20T13:12:00Z"));
            String creator = _coreProperties.Creator;
            ClassicAssert.AreEqual("Paolo Mottadelli", creator);
            String subject = _coreProperties.Subject;
            ClassicAssert.AreEqual("Greetings", subject);
            String title = _coreProperties.Title;
            ClassicAssert.AreEqual("Hello World", title);
        }

        public void TestTransitiveSetters()
        {
            XWPFDocument doc = new XWPFDocument();
            CoreProperties cp = doc.GetProperties().CoreProperties;

            DateTime dateCreated = new DateTime(2010, 6, 15, 10, 0, 0);
            cp.Created = new DateTime(2010, 6, 15, 10, 0, 0);
            ClassicAssert.AreEqual(dateCreated.ToString(), cp.Created.ToString());

            XWPFDocument doc2 = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            doc.Close();
            cp = doc2.GetProperties().CoreProperties;
            DateTime? dt3 = cp.Created;
            ClassicAssert.AreEqual(dateCreated.ToString(), dt3.ToString());

            doc2.Close();
        }
        [Test]
        public void TestGetSetRevision()
        {
            String revision = _coreProperties.Revision;
            ClassicAssert.IsTrue(Int32.Parse(revision) > 1, "Revision number is 1");
            _coreProperties.Revision = "20";
            ClassicAssert.AreEqual("20", _coreProperties.Revision);
            _coreProperties.Revision = "20xx";
            ClassicAssert.AreEqual("20", _coreProperties.Revision);
        }

        [Test]
        public void TestLastModifiedByProperty()
        {
            String lastModifiedBy = _coreProperties.LastModifiedByUser;
            ClassicAssert.AreEqual("Paolo Mottadelli", lastModifiedBy);
            _coreProperties.LastModifiedByUser = "Test User";
            ClassicAssert.AreEqual("Test User", _coreProperties.LastModifiedByUser);
        }


        public static bool DateTimeEqualToUTCString(DateTime? dateTime, String utcString)
        {
            DateTime utcDt = DateTime.SpecifyKind((DateTime)dateTime, DateTimeKind.Utc);
            string dateTimeUtcString = utcDt.ToString("yyyy-MM-ddThh:mm:ssZ");
            return utcString.Equals(dateTimeUtcString);
        }

        [Test]
        public void TestThumbnails()
        {
            POIXMLProperties noThumbProps = sampleNoThumb.GetProperties();

            ClassicAssert.IsNotNull(_props.ThumbnailPart);
            ClassicAssert.IsNull(noThumbProps.ThumbnailPart);

            ClassicAssert.IsNotNull(_props.ThumbnailFilename);
            ClassicAssert.IsNull(noThumbProps.ThumbnailFilename);

            ClassicAssert.IsNotNull(_props.ThumbnailImage);
            ClassicAssert.IsNull(noThumbProps.ThumbnailImage);

            ClassicAssert.AreEqual("thumbnail.jpeg", _props.ThumbnailFilename);


            // Adding / changing
            noThumbProps.SetThumbnail("Testing.png", new ByteArrayInputStream(new byte[1]));
            ClassicAssert.IsNotNull(noThumbProps.ThumbnailPart);
            ClassicAssert.AreEqual("Testing.png", noThumbProps.ThumbnailFilename);
            ClassicAssert.IsNotNull(noThumbProps.ThumbnailImage);
            //ClassicAssert.AreEqual(1, noThumbProps.ThumbnailImage.Available());
            ClassicAssert.AreEqual(1, noThumbProps.ThumbnailImage.Length - noThumbProps.ThumbnailImage.Position);

            noThumbProps.SetThumbnail("Testing2.png", new ByteArrayInputStream(new byte[2]));
            ClassicAssert.IsNotNull(noThumbProps.ThumbnailPart);
            ClassicAssert.AreEqual("Testing.png", noThumbProps.ThumbnailFilename);
            ClassicAssert.IsNotNull(noThumbProps.ThumbnailImage);
            //ClassicAssert.AreEqual(2, noThumbProps.ThumbnailImage.Available());
            ClassicAssert.AreEqual(2, noThumbProps.ThumbnailImage.Length - noThumbProps.ThumbnailImage.Position);
        }


        private static String ZeroPad(long i)
        {
            if (i >= 0 && i <= 9)
            {
                return "0" + i;
            }
            else
            {
                return i.ToString();
            }
        }
    }
}


