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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using NUnit.Framework;
    using System.Threading;
    using NPOI.HSSF.Record;

    /**
     * A Test case for a Test utility class.<br/>
     * Okay, this may seem strange but I need to Test my Test logic.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestSanityChecker
    {
        private static Record INTERFACEHDR = new InterfaceHdrRecord(InterfaceHdrRecord.CODEPAGE);
        private static BoundSheetRecord CreateBoundSheetRec()
        {
            return new BoundSheetRecord("Sheet1");
        }
        [Test]
        public void TestCheckRecordOrder()
        {
            SanityChecker c = new SanityChecker();
            ArrayList records = new ArrayList();
            records.Add(new BOFRecord());
            records.Add(INTERFACEHDR);
            records.Add(CreateBoundSheetRec());
            records.Add(EOFRecord.instance);
            SanityChecker.CheckRecord[] check = new SanityChecker.CheckRecord[]{
				new SanityChecker.CheckRecord(typeof(BOFRecord), '1'),
				new SanityChecker.CheckRecord(typeof(InterfaceHdrRecord), '0'),
				new SanityChecker.CheckRecord(typeof(BoundSheetRecord), 'M'),
				new SanityChecker.CheckRecord(typeof(NameRecord), '*'),
				new SanityChecker.CheckRecord(typeof(EOFRecord), '1'),
		    };
            // Check pass
            c.CheckRecordOrder(records, check);
            records.Insert(2, CreateBoundSheetRec());
            c.CheckRecordOrder(records, check);
            records.RemoveAt(1);	  // optional record missing
            c.CheckRecordOrder(records, check);
            records.Insert(3, new NameRecord());
            records.Insert(3, new NameRecord()); // optional multiple record occurs more than one time
            c.CheckRecordOrder(records, check);

            // Check Assert.Fail
            ConfirmBadRecordOrder(check, new Record[] {
				new BOFRecord(),
				CreateBoundSheetRec(),
				INTERFACEHDR,
				EOFRecord.instance,
		    });

            ConfirmBadRecordOrder(check, new Record[] {
				new BOFRecord(),
				INTERFACEHDR,
				CreateBoundSheetRec(),
				INTERFACEHDR,
				EOFRecord.instance,
		    });

            ConfirmBadRecordOrder(check, new Record[] {
				new BOFRecord(),
				CreateBoundSheetRec(),
				new NameRecord(),
				EOFRecord.instance,
				new NameRecord(),
		    });

            ConfirmBadRecordOrder(check, new Record[] {
				INTERFACEHDR,
				CreateBoundSheetRec(),
				EOFRecord.instance,
		    });

            ConfirmBadRecordOrder(check, new Record[] {
				new BOFRecord(),
				INTERFACEHDR,
				EOFRecord.instance,
		    });

            ConfirmBadRecordOrder(check, new Record[] {
				INTERFACEHDR,
				CreateBoundSheetRec(),
				new BOFRecord(),
				EOFRecord.instance,
		    });

            ConfirmBadRecordOrder(check, new Record[] {
				new BOFRecord(),
				CreateBoundSheetRec(),
				INTERFACEHDR,
				EOFRecord.instance,
		    });
        }

        static SanityChecker.CheckRecord[] check;
        static Record[] recs;
        static void Run()
        {
            try
            {
                SanityChecker c = new SanityChecker();
                IList recs1 = NPOI.Util.Arrays.AsList(recs);
                c.CheckRecordOrder(recs1, check);
            }
            catch (AssertionException)
            {
                // expected during normal Test
                return;
            }
            throw new AssertionException("Did not get Assert.Failure exception as expected");
        }

        private static void ConfirmBadRecordOrder(SanityChecker.CheckRecord[] Check, Record[] recs)
        {
            TestSanityChecker.check = Check;
            TestSanityChecker.recs = recs;

            // ThreadStart ts = new ThreadStart(Run);
            // Thread thread = new Thread(ts);
            // thread.Start();
            // thread.Join();
            Run();
            
        }
    }
}