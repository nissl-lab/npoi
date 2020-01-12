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
    using NUnit.Framework;
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
            Assert.IsNotNull(sampleDoc);
            Assert.IsNotNull(sampleNoThumb);
            _props = sampleDoc.GetProperties();
            _coreProperties = _props.CoreProperties;
            Assert.IsNotNull(_props);
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
            Assert.IsNotNull(props);

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
            Assert.IsTrue(workbook != newWorkbook);


            POIXMLProperties newProps = newWorkbook.GetProperties();
            Assert.IsNotNull(newProps);
            ExtendedProperties newProperties =
                    newProps.ExtendedProperties;

            Assert.AreEqual(application, newProperties.Application);
            Assert.AreEqual(appVersion, newProperties.AppVersion);
        

            CT_ExtendedProperties
                    newCtProps = newProperties.GetUnderlyingProperties();

            Assert.AreEqual(application, newCtProps.Application);
            Assert.AreEqual(appVersion, newCtProps.AppVersion);

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
                Assert.AreEqual("A property with this name already exists in the custom properties", e.Message);
            }
            customProps.AddProperty("test-4", true);

            POIXMLDocument wb2 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack((XSSFWorkbook)wb1);
            wb1.Close();

            CT_CustomProperties ctProps =
                    wb2.GetProperties().CustomProperties.GetUnderlyingProperties();
            Assert.AreEqual(6, ctProps.sizeOfPropertyArray());
            CT_Property p;

            p = ctProps.GetPropertyArray(0);
            Assert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            Assert.AreEqual("test-1", p.name);
            Assert.AreEqual("string val", p.Item.ToString());
            Assert.AreEqual(2, p.pid);

            p = ctProps.GetPropertyArray(1);
            Assert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            Assert.AreEqual("test-2", p.name);
            Assert.AreEqual(1974, p.Item);
            Assert.AreEqual(3, p.pid);

            p = ctProps.GetPropertyArray(2);
            Assert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            Assert.AreEqual("test-3", p.name);
            Assert.AreEqual(36.6, p.Item);
            Assert.AreEqual(4, p.pid);

            p = ctProps.GetPropertyArray(3);
            Assert.AreEqual("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", p.fmtid);
            Assert.AreEqual("test-4", p.name);
            Assert.AreEqual(true, p.Item);
            Assert.AreEqual(5, p.pid);

            p = ctProps.GetPropertyArray(4);
            Assert.AreEqual("Generator", p.name);
            Assert.AreEqual("NPOI", p.Item);
            Assert.AreEqual(6, p.pid);

            //p = ctProps.GetPropertyArray(5);
            //Assert.AreEqual("Generator Version", p.name);
            //Assert.AreEqual("2.0.9", p.Item);
            //Assert.AreEqual(7, p.pid);

            wb2.Close();
        }
        [Ignore("test")]
        public void TestDocumentProperties()
        {
            String category = _coreProperties.Category;
            Assert.AreEqual("test", category);
            String contentStatus = "Draft";
            _coreProperties.ContentStatus = contentStatus;
            Assert.AreEqual("Draft", contentStatus);
            DateTime? Created = _coreProperties.Created;
            // the original file Contains a following value: 2009-07-20T13:12:00Z
            Assert.IsTrue(DateTimeEqualToUTCString(Created, "2009-07-20T13:12:00Z"));
            String creator = _coreProperties.Creator;
            Assert.AreEqual("Paolo Mottadelli", creator);
            String subject = _coreProperties.Subject;
            Assert.AreEqual("Greetings", subject);
            String title = _coreProperties.Title;
            Assert.AreEqual("Hello World", title);
        }

        public void TestTransitiveSetters()
        {
            XWPFDocument doc = new XWPFDocument();
            CoreProperties cp = doc.GetProperties().CoreProperties;

            DateTime dateCreated = new DateTime(2010, 6, 15, 10, 0, 0);
            cp.Created = new DateTime(2010, 6, 15, 10, 0, 0);
            Assert.AreEqual(dateCreated.ToString(), cp.Created.ToString());

            XWPFDocument doc2 = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            doc.Close();
            cp = doc2.GetProperties().CoreProperties;
            DateTime? dt3 = cp.Created;
            Assert.AreEqual(dateCreated.ToString(), dt3.ToString());

            doc2.Close();
        }
        [Test]
        public void TestGetSetRevision()
        {
            String revision = _coreProperties.Revision;
            Assert.IsTrue(Int32.Parse(revision) > 1, "Revision number is 1");
            _coreProperties.Revision = "20";
            Assert.AreEqual("20", _coreProperties.Revision);
            _coreProperties.Revision = "20xx";
            Assert.AreEqual("20", _coreProperties.Revision);
        }

        [Test]
        public void TestLastModifiedByProperty()
        {
            String lastModifiedBy = _coreProperties.LastModifiedByUser;
            Assert.AreEqual("Paolo Mottadelli", lastModifiedBy);
            _coreProperties.LastModifiedByUser = "Test User";
            Assert.AreEqual("Test User", _coreProperties.LastModifiedByUser);
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

            Assert.IsNotNull(_props.ThumbnailPart);
            Assert.IsNull(noThumbProps.ThumbnailPart);

            Assert.IsNotNull(_props.ThumbnailFilename);
            Assert.IsNull(noThumbProps.ThumbnailFilename);

            Assert.IsNotNull(_props.ThumbnailImage);
            Assert.IsNull(noThumbProps.ThumbnailImage);

            Assert.AreEqual("thumbnail.jpeg", _props.ThumbnailFilename);


            // Adding / changing
            noThumbProps.SetThumbnail("Testing.png", new ByteArrayInputStream(new byte[1]));
            Assert.IsNotNull(noThumbProps.ThumbnailPart);
            Assert.AreEqual("Testing.png", noThumbProps.ThumbnailFilename);
            Assert.IsNotNull(noThumbProps.ThumbnailImage);
            //Assert.AreEqual(1, noThumbProps.ThumbnailImage.Available());
            Assert.AreEqual(1, noThumbProps.ThumbnailImage.Length - noThumbProps.ThumbnailImage.Position);

            noThumbProps.SetThumbnail("Testing2.png", new ByteArrayInputStream(new byte[2]));
            Assert.IsNotNull(noThumbProps.ThumbnailPart);
            Assert.AreEqual("Testing.png", noThumbProps.ThumbnailFilename);
            Assert.IsNotNull(noThumbProps.ThumbnailImage);
            //Assert.AreEqual(2, noThumbProps.ThumbnailImage.Available());
            Assert.AreEqual(2, noThumbProps.ThumbnailImage.Length - noThumbProps.ThumbnailImage.Position);
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


