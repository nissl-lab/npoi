/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Extractor
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Extractor;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NUnit.Framework;
    using System.Security.Cryptography;
    [TestFixture]
    public class TestEmbeddedExtractor
    {
        private static  POIDataSamples samples = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void ExtractPDFfromEMF()
        {

            Stream fis = samples.OpenResourceAsStream("Basic_Expense_Template_2011.xls");
            IWorkbook wb = WorkbookFactory.Create(fis);
            fis.Close();

            EmbeddedExtractor ee = new EmbeddedExtractor();
            List<EmbeddedData> edList = new List<EmbeddedData>();
            foreach(ISheet s in wb)
            {
                edList.AddRange(ee.ExtractAll(s));
            }
            wb.Close();

            Assert.AreEqual(2, edList.Count);

            String filename1 = "Sample.pdf";
            EmbeddedData ed0 = edList[0];
            Assert.AreEqual(filename1, ed0.Filename);
            Assert.AreEqual(filename1, ed0.Shape.ShapeName.Trim());
            Assert.AreEqual("uNplB1QpYug+LWappiTh0w==", md5hash(ed0.GetEmbeddedData()));

            String filename2 = "kalastuslupa_jiyjhnj_yuiyuiyuio_uyte_sldfsdfsdf_sfsdfsdf_sfsssfsf_sdfsdfsdfsdf_sdfsdfsdf.pdf";
            EmbeddedData ed1 = edList[1];
            Assert.AreEqual(filename2, ed1.Filename);
            Assert.AreEqual(filename2, ed1.Shape.ShapeName.Trim());
            Assert.AreEqual("QjLuAZ+cd7KbhVz4sj+QdA==", md5hash(ed1.GetEmbeddedData()));
        }

        [Test]
        public void ExtractFromXSSF()
        {

            Stream fis = samples.OpenResourceAsStream("58325_db.xlsx");
            IWorkbook wb = WorkbookFactory.Create(fis);
            fis.Close();

            EmbeddedExtractor ee = new EmbeddedExtractor();
            List<EmbeddedData> edList = new List<EmbeddedData>();
            foreach(ISheet s in wb)
            {
                edList.AddRange(ee.ExtractAll(s));
            }
            wb.Close();

            Assert.AreEqual(4, edList.Count);
            EmbeddedData ed0 = edList[0];
            Assert.AreEqual("Object 1.pdf", ed0.Filename);
            Assert.AreEqual("Object 1", ed0.Shape.ShapeName.Trim());
            Assert.AreEqual("Oyys6UtQU1gbHYBYqA4NFA==", md5hash(ed0.GetEmbeddedData()));

            EmbeddedData ed1 = edList[1];
            Assert.AreEqual("Object 2.pdf", ed1.Filename);
            Assert.AreEqual("Object 2", ed1.Shape.ShapeName.Trim());
            Assert.AreEqual("xLScPUS0XH+5CTZ2A3neNw==", md5hash(ed1.GetEmbeddedData()));

            EmbeddedData ed2 = edList[2];
            Assert.AreEqual("Object 3.pdf", ed2.Filename);
            Assert.AreEqual("Object 3", ed2.Shape.ShapeName.Trim());
            Assert.AreEqual("rX4klZqJAeM5npb54Gi2+Q==", md5hash(ed2.GetEmbeddedData()));

            EmbeddedData ed3 = edList[3];
            Assert.AreEqual("Microsoft_Excel_Worksheet1.xlsx", ed3.Filename);
            Assert.AreEqual("Object 1", ed3.Shape.ShapeName.Trim());
            Assert.AreEqual("4m4N8ji2tjpEGPQuw2YwGA==", md5hash(ed3.GetEmbeddedData()));
        }

        public static String md5hash(byte[] input)
        {
            try
            {
                MD5 md5 = MD5.Create();
                byte[] hash = md5.ComputeHash(input);
                
                return Convert.ToBase64String(hash);
            }
            catch(Exception e)
            {
                // doesn't happen
                throw new RuntimeException(e);
            }
        }


        [Test]
        public void TestNPE()
        {

            HSSFWorkbook wb = HSSF.HSSFTestDataSamples.OpenSampleWorkbook("angelo.edu_content_files_19555-nsse-2011-multiyear-benchmark.xls");
            EmbeddedExtractor ee = new EmbeddedExtractor();

            foreach(ISheet s in wb)
            {
                foreach(EmbeddedData ed in ee.ExtractAll(s))
                {
                    Assert.IsNotNull(ed.Filename);
                    Assert.IsNotNull(ed.GetEmbeddedData());
                    Assert.IsNotNull(ed.Shape);
                }
            }

        }
    }
}

