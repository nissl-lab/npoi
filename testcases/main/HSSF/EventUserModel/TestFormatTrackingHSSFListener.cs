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

namespace TestCases.HSSF.EventUserModel
{
    using System;
    using System.IO;
    using System.Collections;

    using NPOI.HSSF;
    using NPOI.HSSF.EventUserModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;

    using NUnit.Framework;

    [TestFixture]
    public class TestFormatTrackingHSSFListener
    {
        private FormatTrackingHSSFListener listener;
        private MockHSSFListener mockListen;

        private void ProcessFile(String filename)
        {
            HSSFRequest req = new HSSFRequest();
            mockListen = new MockHSSFListener();
            listener = new FormatTrackingHSSFListener(mockListen);
            req.AddListenerForAllRecords(listener);

            HSSFEventFactory factory = new HSSFEventFactory();
            try
            {
                Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(filename);
                POIFSFileSystem fs = new POIFSFileSystem(is1);
                factory.ProcessWorkbookEvents(req, fs);
            }
            catch (IOException)
            {
                throw;
            }
        }
        [Test]
        public void TestFormats()
        {
            ProcessFile("MissingBits.xls");

            Assert.AreEqual("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)", listener.GetFormatString(41));
            Assert.AreEqual("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)", listener.GetFormatString(42));
            Assert.AreEqual("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", listener.GetFormatString(43));
            Assert.AreEqual("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)", listener.GetFormatString(44));
        }

        /**
         * Ensure that all number and formula records can be
         *  turned into strings without problems.
         * For now, we're just looking to Get text back, no
         *  exceptions thrown, but in future we might also
         *  want to check the exact strings!
         */
        [Test]
        public void TestTurnToString()
        {
            String[] files = new String[] { 
				"45365.xls", "45365-2.xls", "MissingBits.xls" 
		};
            for (int k = 0; k < files.Length; k++)
            {
                ProcessFile(files[k]);

                // Check we found our formats
                Assert.IsTrue(listener.NumberOfCustomFormats > 5);
                Assert.IsTrue(listener.NumberOfExtendedFormats > 5);

                // Now check we can turn all the numeric
                //  cells into strings without error
                for (int i = 0; i < mockListen._records.Count; i++)
                {
                    Record r = (Record)mockListen._records[i];
                    CellValueRecordInterface cvr = null;

                    if (r is NumberRecord)
                    {
                        cvr = (CellValueRecordInterface)r;
                    }
                    if (r is FormulaRecord)
                    {
                        cvr = (CellValueRecordInterface)r;
                    }

                    if (cvr != null)
                    {
                        // Should always give us a string 
                        String s = listener.FormatNumberDateCell(cvr);
                        Assert.IsNotNull(s);
                        Assert.IsTrue(s.Length > 0);
                    }
                }

                // TODO - Test some specific format strings
            }
        }

        private class MockHSSFListener : IHSSFListener
        {
            public MockHSSFListener() { }
            internal ArrayList _records = new ArrayList();

            public void ProcessRecord(Record record)
            {
                _records.Add(record);
            }
        }
    }
}