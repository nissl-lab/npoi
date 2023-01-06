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
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Threading;

namespace TestCases.OpenXml4Net.OPC
{

    [TestFixture]
    public class TestPackagingURIHelper
    {

        /**
         * Test relativizePartName() method.
         */
        [Test]
        public void TestRelativizeUri()
        {
            CultureInfo orig = Thread.CurrentThread.CurrentCulture;
            foreach (var ci in System.Globalization.CultureInfo.GetCultures(System.Globalization.CultureTypes.NeutralCultures))
            {
                Thread.CurrentThread.CurrentCulture = ci;
                Uri Uri1 = new Uri("/word/document.xml", UriKind.Relative);
                Uri Uri2 = new Uri("/word/media/image1.gif", UriKind.Relative);
                Uri Uri3 = new Uri("/word/media/image1.gif#Sheet1!A1", UriKind.Relative);
                Uri Uri4 = new Uri("#'My%20Sheet1'!A1", UriKind.Relative);

                // Document to image is down a directory
                Uri retUri1to2 = PackagingUriHelper.RelativizeUri(Uri1, Uri2);
                Assert.AreEqual("media/image1.gif", retUri1to2.OriginalString);
                // Image to document is up a directory
                Uri retUri2to1 = PackagingUriHelper.RelativizeUri(Uri2, Uri1);
                Assert.AreEqual("../document.xml", retUri2to1.OriginalString);

                // Document and CustomXML parts totally different [Julien C.]
                Uri UriCustomXml = new Uri("/customXml/item1.xml", UriKind.RelativeOrAbsolute);

                Uri UriRes = PackagingUriHelper.RelativizeUri(Uri1, UriCustomXml);
                Assert.AreEqual("../customXml/item1.xml", UriRes.ToString());

                // Document to itself is the same place (empty Uri)
                Uri retUri2 = PackagingUriHelper.RelativizeUri(Uri1, Uri1);
                // YK: the line below used to assert empty string which is wrong
                // if source and target are the same they should be relaitivized as the last segment,
                // see Bugzilla 51187
                Assert.AreEqual("document.xml", retUri2.OriginalString);

                // relativization against root
                Uri root = new Uri("/", UriKind.Relative);
                UriRes = PackagingUriHelper.RelativizeUri(root, Uri1);
                Assert.AreEqual("/word/document.xml", UriRes.ToString());

                //Uri compatible with MS Office and OpenOffice: leading slash is Removed
                UriRes = PackagingUriHelper.RelativizeUri(root, Uri1, true);
                Assert.AreEqual("word/document.xml", UriRes.ToString());

                //preserve Uri fragments
                UriRes = PackagingUriHelper.RelativizeUri(Uri1, Uri3, true);
                Assert.AreEqual("media/image1.gif#Sheet1!A1", UriRes.ToString());
                UriRes = PackagingUriHelper.RelativizeUri(root, Uri4, true);
                Assert.AreEqual("#'My%20Sheet1'!A1", UriRes.ToString());
            }
            Thread.CurrentThread.CurrentCulture = orig;
        }

        /**
         * Test CreatePartName(String, y)
         */
        [Test]
        public void TestCreatePartNameRelativeString()
        {
            CultureInfo orig = Thread.CurrentThread.CurrentCulture;
            foreach (var ci in System.Globalization.CultureInfo.GetCultures(System.Globalization.CultureTypes.NeutralCultures))
            {
                Thread.CurrentThread.CurrentCulture = ci;
                PackagePartName partNameToValid = PackagingUriHelper
                    .CreatePartName("/word/media/image1.gif");

                OPCPackage pkg = OPCPackage.Create("DELETEIFEXISTS.docx");
                // Base part
                PackagePartName nameBase = PackagingUriHelper
                        .CreatePartName("/word/document.xml");
                PackagePart partBase = pkg.CreatePart(nameBase, ContentTypes.XML);
                // Relative part name
                PackagePartName relativeName = PackagingUriHelper.CreatePartName(
                        "media/image1.gif", partBase);
                Assert.AreEqual(partNameToValid
                        , relativeName, "The part name must be equal to "
                        + partNameToValid.Name);
                pkg.Revert();
            }
            Thread.CurrentThread.CurrentCulture = orig;
        }

        /**
         * Test CreatePartName(Uri, y)
         */
        [Test]
        public void TestCreatePartNameRelativeUri()
        {
            CultureInfo orig = Thread.CurrentThread.CurrentCulture;
            foreach (var ci in System.Globalization.CultureInfo.GetCultures(System.Globalization.CultureTypes.NeutralCultures))
            {
                Thread.CurrentThread.CurrentCulture = ci;
                PackagePartName partNameToValid = PackagingUriHelper
                    .CreatePartName("/word/media/image1.gif");

                OPCPackage pkg = OPCPackage.Create("DELETEIFEXISTS.docx");
                // Base part
                PackagePartName nameBase = PackagingUriHelper
                        .CreatePartName("/word/document.xml");
                PackagePart partBase = pkg.CreatePart(nameBase, ContentTypes.XML);
                // Relative part name
                PackagePartName relativeName = PackagingUriHelper.CreatePartName(
                        new Uri("media/image1.gif", UriKind.RelativeOrAbsolute), partBase);
                Assert.AreEqual(partNameToValid, relativeName, "The part name must be equal to "
                        + partNameToValid.Name);
                pkg.Revert();
            }
            Thread.CurrentThread.CurrentCulture = orig;
        }
        [Test]
        public void TestCreateUriFromString()
        {
            String[] href = {
                "..\\\\\\cygwin\\home\\yegor\\.vim\\filetype.vim",
                "..\\Program%20Files\\AGEIA%20Technologies\\v2.3.3\\NxCooking.dll",
                "file:///D:\\seva\\1981\\r810102ns.mp3",
                "..\\cygwin\\home\\yegor\\dinom\\%5baccess%5d.2010-10-26.log",
                "#'Instructions (Text)'!B21",
                "javascript://"
         };
            foreach (String s in href)
            {
                try
                {
                    Uri Uri = PackagingUriHelper.ToUri(s);
                }
                catch (UriFormatException)
                {
                    Assert.Fail("Failed to create Uri from " + s);
                }
            }
        }
        [Test]
        public void Test53734()
        {
            Uri uri = PackagingUriHelper.ToUri("javascript://");
            // POI appends a trailing slash tpo avoid "Expected authority at index 13: javascript://"
            // https://issues.apache.org/bugzilla/show_bug.cgi?id=53734
            Assert.AreEqual("javascript:///", uri.ToString());
        }

        [Test]
        public void TestGetFilenameWithoutExtension()
        {
            // Testing fix for th-TH culture see https://github.com/nissl-lab/npoi/issues/944
            String[] href = {
                "..\\\\\\cygwin\\home\\yegor\\.vim\\filetype.vim",
                "..\\Program%20Files\\AGEIA%20Technologies\\v2.3.3\\NxCooking.dll",
                "file:///D:\\seva\\1981\\r810102ns.mp3",
                "..\\cygwin\\home\\yegor\\dinom\\%5baccess%5d.2010-10-26.log",
                "#'Instructions (Text)'!B21",
                "javascript://" };
            String[] fileNameNoExt = {
                "filetype",
                "NxCooking",
                "r810102ns",
                "%5baccess%5d.2010-10-26",
                "",
                "" };

            CultureInfo orig = Thread.CurrentThread.CurrentCulture;
            foreach (var ci in System.Globalization.CultureInfo.GetCultures(System.Globalization.CultureTypes.NeutralCultures))
            {
                Thread.CurrentThread.CurrentCulture = ci;
                for (int idx = 0; idx < href.Length; idx++)
                {
                    try
                    {
                        Uri Uri = PackagingUriHelper.ToUri(href[idx]);
                        String fileName = PackagingUriHelper.GetFilenameWithoutExtension(Uri);
                        Assert.AreEqual(fileNameNoExt[idx], fileName, "GetFilenameWithoutExtension fails with culture : " + ci.Name);
                    }
                    catch (UriFormatException)
                    {
                        Assert.Fail("Failed to create Uri from " + href[idx]);
                    }
                }
            }
            Thread.CurrentThread.CurrentCulture = orig;
        }
    }
}