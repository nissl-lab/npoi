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
using TestCases.OpenXml4Net;
using System.IO;
using NUnit.Framework;
using System;
using System.Collections.Generic;
namespace TestCases.OPC
{
    [TestFixture]
    public class TestListParts
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(TestListParts));

        private Dictionary<PackagePartName, String> expectedValues;

        private Dictionary<PackagePartName, String> values;

        [SetUp]
        public void SetUp()
        {
            values = new Dictionary<PackagePartName, String>();

            // Expected values
            expectedValues = new Dictionary<PackagePartName, String>();
            expectedValues.Add(PackagingUriHelper.CreatePartName("/_rels/.rels"),
                    "application/vnd.openxmlformats-package.relationships+xml");

            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/docProps/app.xml"),
                            "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/docProps/core.xml"),
                    "application/vnd.openxmlformats-package.core-properties+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/word/_rels/document.xml.rels"),
                    "application/vnd.openxmlformats-package.relationships+xml");
            expectedValues
                    .Add(
                            PackagingUriHelper.CreatePartName("/word/document.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/word/fontTable.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/word/media/image1.gif"), "image/gif");
            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/word/settings.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");
            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/word/styles.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/word/theme/theme1.xml"),
                    "application/vnd.openxmlformats-officedocument.theme+xml");
            expectedValues
                    .Add(
                            PackagingUriHelper
                                    .CreatePartName("/word/webSettings.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml");
        }

        /**
         * List all parts of a namespace.
         */
        [Test]
        public void TestListParts1()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("sample.docx");

            OPCPackage p;
            p = OPCPackage.Open(is1);

            foreach (PackagePart part in p.GetParts())
            {
                values.Add(part.PartName, part.ContentType);
                logger.Log(POILogger.DEBUG, part.PartName);
            }

            // Compare expected values with values return by the namespace
            foreach (PackagePartName partName in expectedValues.Keys)
            {
                Assert.IsNotNull(values[partName]);
                Assert.AreEqual(expectedValues[partName], values[partName]);
            }
        }
    }
}



