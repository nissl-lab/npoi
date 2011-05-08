
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
 * Tests the serialization and deserialization of the TextObjectBaseRecord
 * class works correctly.  Test data taken directly from a real
 * Excel file.
 *

 * @author Glen Stampoultzis (glens at apache.org)
 */
public class TestTextObjectBaseRecord
        
{
    byte[] data = new byte[] {
	    0x44, 0x02, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
	    0x00, 0x00, 0x02, 0x00, 0x02, 0x00, 0x00, 0x00,
        0x00, 0x00,
    };

    public TestTextObjectBaseRecord(String name)
    {
        
    }

    public void TestLoad()
            
    {
        TextObjectBaseRecord record = new TextObjectBaseRecord(new TestcaseRecordStream((short)0x1B6, (short)data.Length, data));


//        Assert.AreEqual( (short), record.GetOptions());
        Assert.AreEqual( false, record.IsReserved1() );
        Assert.AreEqual( TextObjectBaseRecord.HORIZONTAL_TEXT_ALIGNMENT_CENTERED, record.GetHorizontalTextAlignment() );
        Assert.AreEqual( TextObjectBaseRecord.VERTICAL_TEXT_ALIGNMENT_JUSTIFY, record.GetVerticalTextAlignment() );
        Assert.AreEqual( 0, record.GetReserved2() );
        Assert.AreEqual( true, record.IsTextLocked() );
        Assert.AreEqual( 0, record.GetReserved3() );
        Assert.AreEqual( TextObjectBaseRecord.TEXT_ORIENTATION_ROT_RIGHT, record.GetTextOrientation());
        Assert.AreEqual( 0, record.GetReserved4());
        Assert.AreEqual( 0, record.GetReserved5());
        Assert.AreEqual( 0, record.GetReserved6());
        Assert.AreEqual( 2, record.GetTextLength());
        Assert.AreEqual( 2, record.GetFormattingRunLength());
        Assert.AreEqual( 0, record.GetReserved7());


        Assert.AreEqual( 22, record.GetRecordSize() );

        record.ValidateSid((short)0x1B6);
    }

    public void TestStore()
    {
        TextObjectBaseRecord record = new TextObjectBaseRecord();



//        record.SetOptions( (short) 0x0000);
        record.SetReserved1( false );
        record.SetHorizontalTextAlignment( TextObjectBaseRecord.HORIZONTAL_TEXT_ALIGNMENT_CENTERED );
        record.SetVerticalTextAlignment( TextObjectBaseRecord.VERTICAL_TEXT_ALIGNMENT_JUSTIFY );
        record.SetReserved2( (short)0 );
        record.SetTextLocked( true );
        record.SetReserved3( (short)0 );
        record.SetTextOrientation( TextObjectBaseRecord.TEXT_ORIENTATION_ROT_RIGHT );
        record.SetReserved4( (short)0 );
        record.SetReserved5( (short)0 );
        record.SetReserved6( (short)0 );
        record.SetTextLength( (short)2 );
        record.SetFormattingRunLength( (short)2 );
        record.SetReserved7( 0 );

        byte [] recordBytes = record.Serialize();
        Assert.AreEqual(recordBytes.Length - 4, data.Length);
        for (int i = 0; i < data.Length; i++)
            Assert.AreEqual("At offSet " + i, data[i], recordBytes[i+4]);
    }
}
