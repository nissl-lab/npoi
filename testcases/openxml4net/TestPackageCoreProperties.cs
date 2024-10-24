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

using NPOI.Util;
using NPOI.OpenXml4Net.OPC;
using System.IO;
using TestCases.OpenXml4Net;
using NUnit.Framework;
using NPOI.SS.Util;
using System;
using NPOI.OpenXmlFormats;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.XSSF.UserModel;
namespace TestCases.OpenXml4Net.OPC
{

    [TestFixture]
    public class TestPackageCoreProperties
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(TestPackageCoreProperties));

        /**
         * Test namespace core properties Getters.
         */
        [Test]
        public void TestGetProperties()
        {
            // Open the namespace
            OPCPackage p = OPCPackage.Open(OpenXml4NetTestDataSamples.OpenSampleStream("TestPackageCoreProperiesGetters.docx"));
            CompareProperties(p);
            p.Revert();

        }

        /**
         * Test namespace core properties Setters.
         */
        [Test]
        public void TestSetProperties()
        {
            String inputPath = OpenXml4NetTestDataSamples.GetSampleFileName("TestPackageCoreProperiesSetters.docx");

            FileInfo outputFile = OpenXml4NetTestDataSamples.GetOutputFile("TestPackageCoreProperiesSettersOUTPUT.docx");

            // Open namespace
            OPCPackage p = OPCPackage.Open(inputPath, PackageAccess.READ_WRITE);

            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
            df.TimeZone = TimeZoneInfo.Utc;
            DateTime dateToInsert = df.Parse("2007-05-12T08:00:00Z");

            SimpleDateFormat msdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.fff'Z'");
            msdf.TimeZone = TimeZoneInfo.Utc;

            PackageProperties props = p.GetPackageProperties();

            //test various date formats
            props.SetCreatedProperty("2007-05-12T08:00:00Z");
            Assert.AreEqual(dateToInsert, props.GetCreatedProperty().Value);

            props.SetCreatedProperty("2007-05-12T08:00:00"); //no Z, assume Z
            Assert.AreEqual(dateToInsert, props.GetCreatedProperty().Value);

            props.SetCreatedProperty("2007-05-12T08:00:00.123Z");//millis
            Assert.AreEqual(msdf.Parse("2007-05-12T08:00:00.123Z"), props.GetCreatedProperty().Value);

            props.SetCreatedProperty("2007-05-12T10:00:00+0200");
            Assert.AreEqual(dateToInsert, props.GetCreatedProperty().Value);

            props.SetCreatedProperty("2007-05-12T10:00:00+02:00");//colon in tz
            Assert.AreEqual(dateToInsert, props.GetCreatedProperty().Value);

            props.SetCreatedProperty("2007-05-12T06:00:00-0200");
            Assert.AreEqual(dateToInsert, props.GetCreatedProperty().Value);

            props.SetCreatedProperty("2015-07-27");
            Assert.AreEqual(msdf.Parse("2015-07-27T00:00:00.000Z"), props.GetCreatedProperty().Value);

            props.SetCreatedProperty("2007-05-12T10:00:00.123+0200");
            Assert.AreEqual(msdf.Parse("2007-05-12T08:00:00.123Z"), props.GetCreatedProperty().Value);

            props.SetCategoryProperty("MyCategory");

            props.SetCategoryProperty("MyCategory");
            props.SetContentStatusProperty("MyContentStatus");
            props.SetContentTypeProperty("MyContentType");
            //props.SetCreatedProperty(new DateTime?(dateToInsert));
            props.SetCreatorProperty("MyCreator");
            props.SetDescriptionProperty("MyDescription");
            props.SetIdentifierProperty("MyIdentifier");
            props.SetKeywordsProperty("MyKeywords");
            props.SetLanguageProperty("MyLanguage");
            props.SetLastModifiedByProperty("Julien Chable");
            props.SetLastPrintedProperty(new Nullable<DateTime>(dateToInsert));
            props.SetModifiedProperty(new Nullable<DateTime>(dateToInsert));
            props.SetRevisionProperty("2");
            props.SetTitleProperty("MyTitle");
            props.SetSubjectProperty("MySubject");
            props.SetVersionProperty("2");
            // Save the package in the output directory
            p.Save(outputFile.FullName);
            p.Revert();


            // Open the newly Created file to check core properties saved values.
            OPCPackage p2 = OPCPackage.Open(outputFile.FullName, PackageAccess.READ);

            CompareProperties(p2);
            p2.Revert();

            outputFile.Delete();

        }

        private void CompareProperties(OPCPackage p)
        {
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
            df.TimeZone = TimeZoneInfo.Utc;
            DateTime expectedDate = df.Parse("2007-05-12T08:00:00Z");

            // Gets the core properties
            PackageProperties props = p.GetPackageProperties();
            Assert.AreEqual("MyCategory", props.GetCategoryProperty());
            Assert.AreEqual("MyContentStatus", props.GetContentStatusProperty()
                    );
            Assert.AreEqual("MyContentType", props.GetContentTypeProperty());
            Assert.AreEqual(expectedDate, props.GetCreatedProperty());
            Assert.AreEqual("MyCreator", props.GetCreatorProperty());
            Assert.AreEqual("MyDescription", props.GetDescriptionProperty());
            Assert.AreEqual("MyIdentifier", props.GetIdentifierProperty());
            Assert.AreEqual("MyKeywords", props.GetKeywordsProperty());
            Assert.AreEqual("MyLanguage", props.GetLanguageProperty());
            Assert.AreEqual("Julien Chable", props.GetLastModifiedByProperty()
                    );
            Assert.AreEqual(expectedDate, props.GetLastPrintedProperty());
            Assert.AreEqual(expectedDate, props.GetModifiedProperty());
            Assert.AreEqual("2", props.GetRevisionProperty());
            Assert.AreEqual("MySubject", props.GetSubjectProperty());
            Assert.AreEqual("MyTitle", props.GetTitleProperty());
            Assert.AreEqual("2", props.GetVersionProperty());
        }
        [Test]
        public void TestCoreProperties_bug51374()
        {
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
            String strDate = "2007-05-12T08:00:00Z";
            DateTime date = DateTime.Parse(strDate).ToUniversalTime();

            OPCPackage pkg = new ZipPackage();
            PackagePropertiesPart props = (PackagePropertiesPart)pkg.GetPackageProperties();

            // Created
            Assert.AreEqual("", props.GetCreatedPropertyString());
            Assert.IsNull(props.GetCreatedProperty());
            props.SetCreatedProperty((String)null);
            Assert.AreEqual("", props.GetCreatedPropertyString());
            Assert.IsNull(props.GetCreatedProperty());
            props.SetCreatedProperty(new Nullable<DateTime>());
            Assert.AreEqual("", props.GetCreatedPropertyString());
            Assert.IsNull(props.GetCreatedProperty());
            props.SetCreatedProperty(new Nullable<DateTime>(date));
            Assert.AreEqual(strDate, props.GetCreatedPropertyString());
            Assert.AreEqual(date, props.GetCreatedProperty());
            props.SetCreatedProperty(strDate);
            Assert.AreEqual(strDate, props.GetCreatedPropertyString());
            Assert.AreEqual(date, props.GetCreatedProperty());

            // lastPrinted
            Assert.AreEqual("", props.GetLastPrintedPropertyString());
            Assert.IsNull(props.GetLastPrintedProperty());
            props.SetLastPrintedProperty((String)null);
            Assert.AreEqual("", props.GetLastPrintedPropertyString());
            Assert.IsNull(props.GetLastPrintedProperty());
            props.SetLastPrintedProperty(new Nullable<DateTime>());
            Assert.AreEqual("", props.GetLastPrintedPropertyString());
            Assert.IsNull(props.GetLastPrintedProperty());
            props.SetLastPrintedProperty(new Nullable<DateTime>(date));
            Assert.AreEqual(strDate, props.GetLastPrintedPropertyString());
            Assert.AreEqual(date, props.GetLastPrintedProperty());
            props.SetLastPrintedProperty(strDate);
            Assert.AreEqual(strDate, props.GetLastPrintedPropertyString());
            Assert.AreEqual(date, props.GetLastPrintedProperty());

            // modified
            Assert.IsNull(props.GetModifiedProperty());
            props.SetModifiedProperty((String)null);
            Assert.IsNull(props.GetModifiedProperty());
            props.SetModifiedProperty(new Nullable<DateTime>());
            Assert.IsNull(props.GetModifiedProperty());
            props.SetModifiedProperty(new Nullable<DateTime>(date));
            Assert.AreEqual(strDate, props.GetModifiedPropertyString());
            Assert.AreEqual(date, props.GetModifiedProperty());
            props.SetModifiedProperty(strDate);
            Assert.AreEqual(strDate, props.GetModifiedPropertyString());
            Assert.AreEqual(date, props.GetModifiedProperty());

            pkg.Close();
        }
        [Test]
        public void TestGetPropertiesLO()
        {
            // Open the namespace
            OPCPackage pkg1 = OPCPackage.Open(OpenXml4NetTestDataSamples.OpenSampleStream("51444.xlsx"));
            PackageProperties props1 = pkg1.GetPackageProperties();
            Assert.AreEqual(null, props1.GetTitleProperty());
            props1.SetTitleProperty("Bug 51444 fixed");
            MemoryStream out1 = new MemoryStream();
            pkg1.Save(out1);
            out1.Close();
            pkg1.Close();

            OPCPackage pkg2 = OPCPackage.Open(new MemoryStream(out1.ToArray()));
            PackageProperties props2 = pkg2.GetPackageProperties();
            props2.SetTitleProperty("Bug 51444 fixed");
            pkg2.Close();
        }
        [Test]
        public void TestEntitiesInCoreProps_56164()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("CorePropertiesHasEntities.ooxml");
            OPCPackage p = OPCPackage.Open(is1);
            is1.Close();

            // Should have 3 root relationships
            bool foundDocRel = false, foundCorePropRel = false, foundExtPropRel = false;
            foreach (PackageRelationship pr in p.Relationships)
            {
                if (pr.RelationshipType.Equals(PackageRelationshipTypes.CORE_DOCUMENT))
                    foundDocRel = true;
                if (pr.RelationshipType.Equals(PackageRelationshipTypes.CORE_PROPERTIES))
                    foundCorePropRel = true;
                if (pr.RelationshipType.Equals(PackageRelationshipTypes.EXTENDED_PROPERTIES))
                    foundExtPropRel = true;
            }
            Assert.IsTrue(foundDocRel, "Core/Doc Relationship not found in " + p.Relationships);
            Assert.IsTrue(foundCorePropRel, "Core Props Relationship not found in " + p.Relationships);
            Assert.IsTrue(foundExtPropRel, "Ext Props Relationship not found in " + p.Relationships);

            // Get the Core Properties
            PackagePropertiesPart props = (PackagePropertiesPart)p.GetPackageProperties();

            // Check
            Assert.AreEqual("Stefan Kopf", props.GetCreatorProperty());

            p.Close();
        }

        [Test]
        public void TestListOfCustomProperties()
        {
            FileInfo inp = POIDataSamples.GetSpreadSheetInstance().GetFileInfo("ExcelWithAttachments.xlsm");
            OPCPackage pkg = OPCPackage.Open(inp, PackageAccess.READ);
            XSSFWorkbook wb = new XSSFWorkbook(pkg);

            Assert.IsNotNull(wb.GetProperties());
            Assert.IsNotNull(wb.GetProperties().CustomProperties);

            foreach (CT_Property prop in wb.GetProperties().CustomProperties.GetUnderlyingProperties().GetPropertyList())
            {
                Assert.IsNotNull(prop);
            }

            wb.Close();
            pkg.Close();
        }


        [Test]
        public void TestAlternateCorePropertyTimezones()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("OPCCompliance_CoreProperties_AlternateTimezones.docx");
            OPCPackage pkg = OPCPackage.Open(is1);
            PackagePropertiesPart props = (PackagePropertiesPart)pkg.GetPackageProperties();
            is1.Close();

            // We need predictable dates for testing!
            //SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.fff'Z'"); //use fff for millisecond.
            df.TimeZone = TimeZoneInfo.Utc;
            // Check text properties first
            Assert.AreEqual("Lorem Ipsum", props.GetTitleProperty());
            Assert.AreEqual("Apache POI", props.GetCreatorProperty());
            
            // Created at has a +3 timezone and milliseconds
            //   2006-10-13T18:06:00.123+03:00
            // = 2006-10-13T15:06:00.123+00:00
            Assert.AreEqual("2006-10-13T15:06:00Z", props.GetCreatedPropertyString());
            Assert.AreEqual("2006-10-13T15:06:00.123Z", df.Format(props.GetCreatedProperty()));

            // Modified at has a -13 timezone but no milliseconds
            //   2007-06-20T07:59:00-13:00
            // = 2007-06-20T20:59:00-13:00
            Assert.AreEqual("2007-06-20T20:59:00Z", props.GetModifiedPropertyString());
            Assert.AreEqual("2007-06-20T20:59:00.000Z", df.Format(props.GetModifiedProperty()));


            // Ensure we can change them with other timezones and still read back OK
            props.SetCreatedProperty("2007-06-20T20:57:00+13:00");
            props.SetModifiedProperty("2007-06-20T20:59:00.123-13:00");

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            pkg.Save(baos);
            pkg = OPCPackage.Open(new ByteArrayInputStream(baos.ToByteArray()));

            // Check text properties first - should be unchanged
            Assert.AreEqual("Lorem Ipsum", props.GetTitleProperty());
            Assert.AreEqual("Apache POI", props.GetCreatorProperty());

            // Check the updated times
            //   2007-06-20T20:57:00+13:00
            // = 2007-06-20T07:57:00Z
            Assert.AreEqual("2007-06-20T07:57:00.000Z", df.Format(props.GetCreatedProperty().Value));

            //   2007-06-20T20:59:00.123-13:00
            // = 2007-06-21T09:59:00.123Z
            Assert.AreEqual("2007-06-21T09:59:00.123Z", df.Format(props.GetModifiedProperty().Value));
        }


    }
}




