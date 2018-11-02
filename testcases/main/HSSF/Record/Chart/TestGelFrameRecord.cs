
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
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Chart;
using NUnit.Framework;
using NPOI.Util;

namespace TestCases.HSSF.Record.Chart
{
    /// <summary>
    /// 
    /// </summary>
    /// <remarks>
    /// author: Antony liu (antony.apollo at gmail.com)
    /// </remarks>
    [TestFixture]
    class TestGelFrameRecord
    {
        //No Test Data
        byte[] data = HexRead.ReadFromString(@"E3 01 0B F0 B4 00 00 00 80 01 00 00
 00 00 81 01 9B BB 59 02 82 01 00 00 01 00 83 01 
 FF FF FF 02 84 01 00 00 01 00 85 01 F4 00 00 10 
 86 C1 00 00 00 00 87 C1 00 00 00 00 88 01 00 00 
 00 00 89 01 00 00 00 00 8A 01 00 00 00 00 8B 01 
 00 00 00 00 8C 01 00 00 00 00 8D 01 00 00 00 00 
 8E 01 00 00 00 00 8F 01 00 00 00 00 90 01 00 00 
 00 00 91 01 00 00 00 00 92 01 00 00 00 00 93 01 
 00 00 00 00 94 01 00 00 00 00 95 01 00 00 00 00 
 96 01 00 00 00 00 97 C1 00 00 00 00 98 01 00 00 
 00 00 99 01 00 00 00 00 9A 01 00 00 00 00 9B 01 
 00 00 00 00 9C 01 03 00 00 40 BF 01 1C 00 1F 00 
 B3 00 22 F1 42 00 00 00 9E 01 FF FF FF FF 9F 01 
 FF FF FF FF A0 01 00 00 00 20 A1 C1 00 00 00 00 
 A2 01 FF FF FF FF A3 01 FF FF FF FF A4 01 00 00 
 00 20 A5 C1 00 00 00 00 A6 01 FF FF FF FF A7 01 
 FF FF FF FF BF 01 00 00 60 00");
        [Test]
        public void TestLoad()
        {
            GelFrameRecord record = new GelFrameRecord(TestcaseRecordInputStream.Create((short)0x1066, data));
            //Assert.AreEqual(0xD, record.Options);
        }
    }
}
