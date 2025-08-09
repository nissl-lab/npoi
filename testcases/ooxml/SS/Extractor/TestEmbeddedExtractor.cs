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
using System.Collections.Generic;
using System.IO;

namespace TestCases.SS.Extractor
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Extractor;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System.Security.Cryptography;
    using TestCases.HSSF;

    [TestFixture]
    public class TestEmbeddedExtractor
    {
        private static POIDataSamples samples = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void ExtractPDFfromEMF()
        {
            Stream fis = samples.OpenResourceAsStream("Basic_Expense_Template_2011.xls");
            IWorkbook wb = WorkbookFactory.Create(fis);
            fis.Close();

            EmbeddedExtractor ee = new EmbeddedExtractor();
            List<EmbeddedData> edList = new List<EmbeddedData>();
            foreach (ISheet s in wb) {
                edList.AddRange(ee.ExtractAll(s));
            }
            wb.Close();

            ClassicAssert.AreEqual(2, edList.Count);

            string filename1 = "Sample.pdf";
            EmbeddedData ed0 = edList[0];
            ClassicAssert.AreEqual(filename1, ed0.Filename);
            ClassicAssert.AreEqual(filename1, ed0.Shape.ShapeName.Trim());
            ClassicAssert.AreEqual("uNplB1QpYug+LWappiTh0w==", md5hash(ed0.GetEmbeddedData()));

            string filename2 = "kalastuslupa_jiyjhnj_yuiyuiyuio_uyte_sldfsdfsdf_sfsdfsdf_sfsssfsf_sdfsdfsdfsdf_sdfsdfsdf.pdf";
            EmbeddedData ed1 = edList[1];
            ClassicAssert.AreEqual(filename2, ed1.Filename);
            ClassicAssert.AreEqual(filename2, ed1.Shape.ShapeName.Trim());
            ClassicAssert.AreEqual("QjLuAZ+cd7KbhVz4sj+QdA==", md5hash(ed1.GetEmbeddedData()));
        }

        [Test]
        public void ExtractFromXSSF()
        {
            Stream fis = samples.OpenResourceAsStream("58325_db.xlsx");
            IWorkbook wb = WorkbookFactory.Create(fis);
            fis.Close();

            EmbeddedExtractor ee = new EmbeddedExtractor();
            List<EmbeddedData> edList = new List<EmbeddedData>();
            foreach (ISheet s in wb) {
                edList.AddRange(ee.ExtractAll(s));
            }
            wb.Close();

            ClassicAssert.AreEqual(4, edList.Count);
            EmbeddedData ed0 = edList[0];
            ClassicAssert.AreEqual("Object 1.pdf", ed0.Filename);
            ClassicAssert.AreEqual("Object 1", ed0.Shape.ShapeName.Trim());
            ClassicAssert.AreEqual("Oyys6UtQU1gbHYBYqA4NFA==", md5hash(ed0.GetEmbeddedData()));

            EmbeddedData ed1 = edList[1];
            ClassicAssert.AreEqual("Object 2.pdf", ed1.Filename);
            ClassicAssert.AreEqual("Object 2", ed1.Shape.ShapeName.Trim());
            ClassicAssert.AreEqual("xLScPUS0XH+5CTZ2A3neNw==", md5hash(ed1.GetEmbeddedData()));

            EmbeddedData ed2 = edList[2];
            ClassicAssert.AreEqual("Object 3.pdf", ed2.Filename);
            ClassicAssert.AreEqual("Object 3", ed2.Shape.ShapeName.Trim());
            ClassicAssert.AreEqual("rX4klZqJAeM5npb54Gi2+Q==", md5hash(ed2.GetEmbeddedData()));

            EmbeddedData ed3 = edList[3];
            ClassicAssert.AreEqual("Microsoft_Excel_Worksheet1.xlsx", ed3.Filename);
            ClassicAssert.AreEqual("Object 1", ed3.Shape.ShapeName.Trim());
            ClassicAssert.AreEqual("4m4N8ji2tjpEGPQuw2YwGA==", md5hash(ed3.GetEmbeddedData()));
        }

        public static string md5hash(byte[] input) {
            
            MD5 md5 = MD5.Create();

            byte[] hash = md5.ComputeHash(input);
            return Convert.ToBase64String(hash);
        }


        [Test]
        public void TestNPE()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("angelo.edu_content_files_19555-nsse-2011-multiyear-benchmark.xls");
            EmbeddedExtractor ee = new EmbeddedExtractor();

            foreach (ISheet s in wb)
            {
                foreach (EmbeddedData ed in ee.ExtractAll(s))
                {
                    ClassicAssert.IsNotNull(ed.Filename);
                    ClassicAssert.IsNotNull(ed.GetEmbeddedData());
                    ClassicAssert.IsNotNull(ed.Shape);
                }
            }

        }
    }
}


