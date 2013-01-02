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

namespace TestCases.HPSF.Basic
{
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using System;
    using NPOI.SS.UserModel;

    /**
     * Tests various bugs have been fixed
     */
    [TestFixture]
    public class TestHPSFBugs
    {
        /**
         * Ensure that we can create a new HSSF Workbook,
         *  then add some properties to it, save +
         *  reload, and still access & change them.
         */
        [Test]
        public void Test48832()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // Starts empty
            Assert.IsNull(wb.DocumentSummaryInformation);
            Assert.IsNull(wb.SummaryInformation);

            // Add new properties
            wb.CreateInformationProperties();

            Assert.IsNotNull(wb.DocumentSummaryInformation);
            Assert.IsNotNull(wb.SummaryInformation);

            // Set Initial values
            wb.SummaryInformation.Author = (/*setter*/"Apache POI");
            wb.SummaryInformation.Keywords = (/*setter*/"Testing POI");
            wb.SummaryInformation.CreateDateTime = DateUtil.GetJavaDate(12345);

            wb.DocumentSummaryInformation.Company = (/*setter*/"Apache");


            // Save and reload
            MemoryStream baos = new MemoryStream();
            wb.Write(baos);
            MemoryStream bais =
               new MemoryStream(baos.ToArray());
            wb = new HSSFWorkbook(bais);


            // Ensure Changes were taken
            Assert.IsNotNull(wb.DocumentSummaryInformation);
            Assert.IsNotNull(wb.SummaryInformation);

            Assert.AreEqual("Apache POI", wb.SummaryInformation.Author);
            Assert.AreEqual("Testing POI", wb.SummaryInformation.Keywords);
            Assert.AreEqual(12345, DateUtil.GetExcelDate(wb.SummaryInformation.CreateDateTime.Value));
            Assert.AreEqual("Apache", wb.DocumentSummaryInformation.Company);


            // Set some more, save + reload
            wb.SummaryInformation.Comments = (/*setter*/"Resaved");

            baos = new MemoryStream();
            wb.Write(baos);
            bais = new MemoryStream(baos.ToArray());
            wb = new HSSFWorkbook(bais);

            // Check again
            Assert.IsNotNull(wb.DocumentSummaryInformation);
            Assert.IsNotNull(wb.SummaryInformation);

            Assert.AreEqual("Apache POI", wb.SummaryInformation.Author);
            Assert.AreEqual("Testing POI", wb.SummaryInformation.Keywords);
            Assert.AreEqual("Resaved", wb.SummaryInformation.Comments);
            Assert.AreEqual(12345, DateUtil.GetExcelDate(wb.SummaryInformation.CreateDateTime.Value));
            Assert.AreEqual("Apache", wb.DocumentSummaryInformation.Company);
        }
    }

}