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
using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using NPOI.HSSF.UserModel;
using TestCases.HSSF.UserModel;

namespace TestCases.HSSF.Record.Aggregates.Chart
{
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.Model;
    using NPOI.Util;
    using NPOI.HSSF.Record.Aggregates.Chart;
    using NPOI.HSSF.Record.Chart;
    [TestFixture]
    public class TestChartSheetAggregate
    {
        [Test]
        public void TestStartBlock_EndBlock_Write()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("chartdemo.xls");
            Record[] sheetRecs = RecordInspector.GetRecords(wb.GetSheetAt(0), 0);

            RecordStream rs = new RecordStream(Arrays.AsList(sheetRecs), 0);

            rs.FindChartSubStream();
            int pos = rs.GetCountRead();

            ChartSheetAggregate csAgg = new ChartSheetAggregate(rs, null);
            RecordInspector.RecordCollector rv = new RecordInspector.RecordCollector();
            csAgg.VisitContainedRecords(rv);
            Record[] outRecs = rv.Records;
            for (int i = 0; i < outRecs.Length; i++)
            {
                Assert.AreEqual(sheetRecs[pos + i].GetType(), outRecs[i].GetType());
            }
        }
    }
    
}
