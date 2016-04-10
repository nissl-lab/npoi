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
using TestCases.OpenXml4Net;
using System;
using System.IO;
using NUnit.Framework;
namespace TestCases.OPC
{

    /**
     * Test the Addition of thumbnail in a namespace.
     * 
     * @author Julien Chable
     */
    [TestFixture]
    public class TestPackageThumbnail
    {

        /**
         * Test namespace AddThumbnail() method.
         */
        [Test]
        public void TestSetProperties()
        {
            String inputPath = OpenXml4NetTestDataSamples.GetSampleFileName("TestPackageThumbnail.docx");

            String imagePath = OpenXml4NetTestDataSamples.GetSampleFileName("thumbnail.jpg");

            FileInfo outputFile = OpenXml4NetTestDataSamples.GetOutputFile("TestPackageThumbnailOUTPUT.docx");

            // Open namespace
            using (Stream inputfile = File.OpenRead(inputPath))
            {
                OPCPackage p = OPCPackage.Open(inputfile);
                try
                {
                    using (FileStream fs = outputFile.OpenWrite())
                    {
                        p.AddThumbnail(imagePath);
                        // Save the namespace in the output directory
                        p.Save(fs);
                        p.Close();
                    }

                    // Open the newly created file to check core properties saved values.
                    OPCPackage p2 = OPCPackage.Open(outputFile.Name, PackageAccess.READ);
                    try
                    {
                        if (p2.GetRelationshipsByType(PackageRelationshipTypes.THUMBNAIL)
                                .Size == 0)
                            Assert.Fail("Thumbnail not added to the namespace !");
                        p2.Revert();

                    }
                    finally
                    {
                        p2.Revert();
                        p2.Close();
                    }
                }
                finally
                {
                    File.Delete(outputFile.Name);
                }
            }
        }
    }


}

