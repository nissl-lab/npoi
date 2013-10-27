/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record.Chart
{
    using System;
    using System.Collections;
    using System.IO;
    using NUnit.Framework;

    using TestCases.HSSF;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.EventUserModel;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI.HSSF.Record.Chart;
    /**
     * 
     */
    [TestFixture]
    public class TestAlRunsRecord
    {
        [Test]
        public void TestRecord()
        {
            POIFSFileSystem fs = new POIFSFileSystem(
                    HSSFTestDataSamples.OpenSampleFileStream("WithFormattedGraphTitle.xls"));

            // Check we can Open the file via usermodel
            HSSFWorkbook hssf = new HSSFWorkbook(fs);

            // Now process it through eventusermodel, and
            //  look out for the title records
            ChartTitleFormatRecordGrabber grabber = new ChartTitleFormatRecordGrabber();
            Stream din = fs.CreateDocumentInputStream("Workbook");
            HSSFRequest req = new HSSFRequest();
            req.AddListenerForAllRecords(grabber);
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.ProcessEvents(req, din);
            din.Close();

            // Should've found one
            Assert.AreEqual(1, grabber.chartTitleFormatRecords.Count);
            // And it should be of something interesting
            AlRunsRecord r =
                (AlRunsRecord)grabber.chartTitleFormatRecords[0];
            Assert.AreEqual(3, r.GetFormatCount());
        }

        private class ChartTitleFormatRecordGrabber : IHSSFListener
        {
            public IList chartTitleFormatRecords;

            public ChartTitleFormatRecordGrabber()
            {
                chartTitleFormatRecords = new ArrayList();
            }

            public void ProcessRecord(Record record)
            {
                if (record is AlRunsRecord)
                {
                    chartTitleFormatRecords.Add(
                            record
                    );
                }
            }

        }
    }
}