
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
        

namespace NPOI.HSSF.Record
{
    using System.Text;
    using NPOI.Util;
    using System;


/**
 * End Of File record.
 * 
 * Description:  Marks the end of records belonging to a particular object in the
 *               HSSF File
 * REFERENCE:  PG 307 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
 * @author Andrew C. Oliver (acoliver at apache dot org)
 * @author Jason Height (jheight at chariot dot net dot au)
 * @version 2.0-pre
 */

public class EOFRecord: StandardRecord
{
    public const short sid = 0x0A;
    public const int ENCODED_SIZE = 4;
    public static readonly EOFRecord instance = new EOFRecord();

    public EOFRecord()
    {
    }

    /**
     * Constructs a EOFRecord record and Sets its fields appropriately.
     * @param in the RecordInputstream to Read the record from
     */

    public EOFRecord(RecordInputStream in1)
    {
        
    }

    public override String ToString()
    {
        StringBuilder buffer = new StringBuilder();

        buffer.Append("[EOF]\n");
        buffer.Append("[/EOF]\n");
        return buffer.ToString();
    }

    public override void Serialize(ILittleEndianOutput out1)
    {
    }

    protected override int DataSize
    {
        get { return ENCODED_SIZE-4; }
    }

    public override short Sid
    {
        get{return sid;}
    }

    public override Object Clone() {
      EOFRecord rec = new EOFRecord();
      return rec;
    }
}
}