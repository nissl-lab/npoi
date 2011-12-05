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

namespace NPOI
{

    using NPOI.XSSF.XSSFTestDataSamples;
    using NPOI.XSSF.usermodel.XSSFWorkbook;
    using NPOI.XWPF.XWPFTestDataSamples;
    using NPOI.XWPF.usermodel.XWPFDocument;
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using OpenXmlFormats;

    /**
     * Test Setting extended and custom OOXML properties
     */
    public class TestPOIXMLProperties
    {
        private POIXMLProperties _props;
        private CoreProperties _coreProperties;

        public void SetUp()
        {
            XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("documentProperties.docx");
            _props = sampleDoc.GetProperties();
            _coreProperties = _props.GetCoreProperties();
            assertNotNull(_props);
        }

        public void testWorkbookExtendedProperties()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            POIXMLProperties props = workbook.GetProperties();
            Assert.IsNotNull(props);

            NPOI.POIXMLPROPERTIES.ExtendedProperties properties =
                    props.GetExtendedProperties();

            CT_Properties
                    ctProps = properties.GetUnderlyingProperties();


            String appVersion = "3.5 beta";
            String application = "POI";

            ctProps.Application=(application);
            ctProps.AppVersion=(appVersion);

            ctProps = null;
            properties = null;
            props = null;

            XSSFWorkbook newWorkbook =
                    XSSFTestDataSamples.WriteOutAndReadBack(workbook);

            Assert.IsTrue(workbook != newWorkbook);


            POIXMLProperties newProps = newWorkbook.GetProperties();
            Assert.IsNotNull(newProps);
            NPOI.POIXMLPROPERTIES.ExtendedProperties newProperties =
                    newProps.GetExtendedProperties();

            CT_Properties
                    newCtProps = newProperties.GetUnderlyingProperties();

            Assert.AreEqual(application, newCtProps.Application);
            Assert.AreEqual(appVersion, newCtProps.AppVersion);


        }


        /**
         * Test usermodel API for Setting custom properties
         */
        public void TestCustomProperties()
        {
            POIXMLDocument wb = new XSSFWorkbook();

            POIXMLProperties.CustomProperties customProps = wb.GetProperties().GetCustomProperties();
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
                Assert.AreEqual("A property with this name already exists in the custom properties", e.GetMessage());
            }
            customProps.AddProperty("test-4", true);

            wb = XSSFTestDataSamples.WriteOutAndReadBack((XSSFWorkbook)wb);
            CT_Properties ctProps =
                    wb.GetProperties().GetCustomProperties().GetUnderlyingProperties();
            Assert.AreEqual(4, ctProps.sizeOfPropertyArray());
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
        }

        public void TestDocumentProperties()
        {
            String category = _coreProperties.GetCategory();
            Assert.AreEqual("test", category);
            String contentStatus = "Draft";
            _coreProperties.SetContentStatus(contentStatus);
            Assert.AreEqual("Draft", contentStatus);
            DateTime Created = _coreProperties.GetCreated();
            // the original file Contains a following value: 2009-07-20T13:12:00Z
            Assert.IsTrue(DateTimeEqualToUTCString(Created, "2009-07-20T13:12:00Z"));
            String creator = _coreProperties.GetCreator();
            Assert.AreEqual("Paolo Mottadelli", creator);
            String subject = _coreProperties.GetSubject();
            Assert.AreEqual("Greetings", subject);
            String title = _coreProperties.GetTitle();
            Assert.AreEqual("Hello World", title);
        }

        public void TestTransitiveSetters()
        {
            XWPFDocument doc = new XWPFDocument();
            NPOI.POIXMLProperties.CoreProperties cp = doc.GetProperties().GetCoreProperties();

            DateTime dateCreated = new GregorianCalendar(2010, 6, 15, 10, 0, 0).GetTime();
            cp.SetCreated(new Nullable<Date>(dateCreated));
            Assert.AreEqual(dateCreated.ToString(), cp.GetCreated().ToString());

            doc = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            cp = doc.GetProperties().GetCoreProperties();
            DateTime dt3 = cp.GetCreated();
            Assert.AreEqual(dateCreated.ToString(), dt3.ToString());

        }

        public void TestGetSetRevision()
        {
            String revision = _coreProperties.GetRevision();
            Assert.IsTrue(Int32.Parse(revision) > 1, "Revision number is 1");
            _coreProperties.SetRevision("20");
            Assert.AreEqual("20", _coreProperties.GetRevision());
            _coreProperties.SetRevision("20xx");
            Assert.AreEqual("20", _coreProperties.GetRevision());
        }

        public static bool DateTimeEqualToUTCString(DateTime dateTime, String utcString)
        {
            Calendar utcCalendar = Calendar.GetInstance(TimeZone.GetTimeZone("UTC"), Locale.UK);
            utcCalendar.SetTimeInMillis(dateTime.GetTime());
            String dateTimeUtcString = utcCalendar.Get(Calendar.YEAR) + "-" +
                   ZeroPad((utcCalendar.Get(Calendar.MONTH) + 1)) + "-" +
                   ZeroPad(utcCalendar.Get(Calendar.DAY_OF_MONTH)) + "T" +
                   ZeroPad(utcCalendar.Get(Calendar.HOUR_OF_DAY)) + ":" +
                   ZeroPad(utcCalendar.Get(Calendar.MINUTE)) + ":" +
                   ZeroPad(utcCalendar.Get(Calendar.SECOND)) + "Z";


            return utcString.Equals(dateTimeUtcString);
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
}


