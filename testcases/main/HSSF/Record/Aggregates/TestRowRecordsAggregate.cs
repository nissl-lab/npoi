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

namespace TestCases.HSSF.Record.Aggregates
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;

    /**
     * 
     */
    [TestClass]
    public class TestRowRecordsAggregate
    {
        [TestMethod]
        public void TestRowGet()
        {
            RowRecordsAggregate rra = new RowRecordsAggregate();
            RowRecord rr = new RowRecord(4);
            rra.InsertRow(rr);
            rra.InsertRow(new RowRecord(1));

            RowRecord rr1 = rra.GetRow(4);

            Assert.IsNotNull(rr1);
            Assert.AreEqual(4, rr1.RowNumber, "Row number is1 1");
            Assert.IsTrue(rr1 == rr, "Row record retrieved is1 identical");
        }
    }
}