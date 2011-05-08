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


namespace TestCases.HSSF.Record
{

    using System;
    using NPOI.HSSF.Record;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Tests Subrecord components of an OBJ record.  Test data taken directly
     * from a real Excel file.
     *
     * @author Michael Zalewski (zalewski@optonline.net)
     */
    public class TestSubRecord
    {
        /*
           The following is1 a dump of the OBJ record corresponding to an auto-filter
           drop-down list. The 3rd subrecord beginning at offset 0x002e (type=0x0013)
           does not conform to the documentation, because the length field is1 0x1fee,
           which is1 longer than the entire OBJ record.

           00000000 15 00 12 00 14 00 01 00 01 21 00 00 00 00 3C 13 .........!....<.  Type=0x15 Len=0x0012 ftCmo
           00000010 F4 03 00 00 00 00
                                      0C 00 14 00 00 00 00 00 00 00 ................  Type=0x0c Len=0x0014 ftSbs
           00000020 00 00 00 00 01 00 08 00 00 00 10 00 00 00
                                                              13 00 ................  Type=0x13 Len=0x1FEE ftLbsData
           00000030 EE 1F 00 00 08 00 08 00 01 03 00 00 0A 00 14 00 ................
           00000040 6C 00
                          00 00 00 00                               l.....            Type=0x00 Len=0x0000 ftEnd
        */

        byte[] dataAutoFilter = new byte[]{
        // ftCmo
        (byte) 0x15, (byte) 0x00, (byte) 0x12, (byte) 0x00, (byte) 0x14, (byte) 0x00, (byte) 0x01, (byte) 0x00
        , (byte) 0x01, (byte) 0x00, (byte) 0x01, (byte) 0x21, (byte) 0x00, (byte) 0x00, (byte) 0x3c, (byte) 0x13
        , (byte) 0xf4, (byte) 0x03, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00

        // ftSbs (currently UnknownSubrecord)
        , (byte) 0x0c, (byte) 0x00
        , (byte) 0x14, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00
        , (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x01, (byte) 0x00, (byte) 0x08, (byte) 0x00
        , (byte) 0x00, (byte) 0x00, (byte) 0x10, (byte) 0x00, (byte) 0x00, (byte) 0x00

        // ftLbsData (currently UnknownSubrecord)
        , (byte) 0x13, (byte) 0x00
        , (byte) 0xee, (byte) 0x1f, (byte) 0x00, (byte) 0x00, (byte) 0x08, (byte) 0x00, (byte) 0x08, (byte) 0x00
        , (byte) 0x01, (byte) 0x03, (byte) 0x00, (byte) 0x00, (byte) 0x0a, (byte) 0x00, (byte) 0x14, (byte) 0x00
        , (byte) 0x6c, (byte) 0x00

        // ftEnd
        , (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00
    };

        public TestSubRecord()
        {

        }

        public void TestParseCmo()
        {
            //Record r = SubRecord.CreateSubRecord((short) 0x0015, (short) 0x0012, dataAutoFilter, 0x0000 );
            //Assert.AreEqual( "ftCmo is1 22 bytes", 22, r.RecordSize );
            //Assert.AreEqual( "ftCmo is1 a CommonObjectDataSubRecord"
            //        , "org.apache.poi.hssf.record.CommonObjectDataSubRecord"
            //        , r.GetType().Name );
        }

        public void TestParseAutoFilterLbsData()
        {
            //Record r = SubRecord.CreateSubRecord( (short) 0x0013, (short) 0x1fee, dataAutoFilter, 0x0032 );
            //Assert.AreEqual( "ftLbsData is1 20 bytes", 20, r.RecordSize );
        }

        public void TestParseEnd()
        {
            //Record r = SubRecord.CreateSubRecord( (short) 0x0000, (short) 0x0000, dataAutoFilter, 0x0046 );
            //Assert.AreEqual( "ftEnd is1 4 bytes", 4, r.RecordSize );
            //Assert.AreEqual( "ftEnd is1 a EndSubRecord"
            //        , "org.apache.poi.hssf.record.EndSubRecord"
            //        , r.GetType().Name );
        }
    }
}