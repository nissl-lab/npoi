
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
using NPOI.Util;

using System;
    using System.Collections;
    using System.Text;

/**
 * Title:        Extended Static String Table
 * Description: This record Is used for a quick Lookup into the SST record. This
 *              record breaks the SST table into a Set of buckets. The offsets
 *              to these buckets within the SST record are kept as well as the
 *              position relative to the start of the SST record.
 * REFERENCE:  PG 313 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
 * @author Andrew C. Oliver (acoliver at apache dot org)
 * @author Jason Height (jheight at apache dot org)
 * @version 2.0-pre
 * @see org.apache.poi.hssf.record.ExtSSTInfoSubRecord
 */

public class ExtSSTRecord: Record
{
    public const short DEFAULT_BUCKET_SIZE = 8;
    //Cant seem to Find this documented but from the biffviewer it Is clear that
    //Excel only records the indexes for the first 128 buckets.
    public const int MAX_BUCKETS = 128;
    public const short sid = 0xff;
    private short             field_1_strings_per_bucket = DEFAULT_BUCKET_SIZE;
    private ArrayList         field_2_sst_info;


    public ExtSSTRecord()
    {
        field_2_sst_info = new ArrayList();
    }

    /**
     * Constructs a EOFRecord record and Sets its fields appropriately.
     * @param in the RecordInputstream to Read the record from
     */

    public ExtSSTRecord(RecordInputStream in1)
    {
        field_2_sst_info = new ArrayList();
        field_1_strings_per_bucket = in1.ReadShort();
        while (in1.Remaining > 0)
        {
            ExtSSTInfoSubRecord rec = new ExtSSTInfoSubRecord(in1);

            field_2_sst_info.Add(rec);
        }
    }

    public IEnumerator GetEnumerator()
    {
        return field_2_sst_info.GetEnumerator();
    }

    public void AddInfoRecord(ExtSSTInfoSubRecord rec)
    {
        field_2_sst_info.Add(rec);
    }

    public short NumStringsPerBucket
    {
        get
        {
            return field_1_strings_per_bucket;
        }
        set { field_1_strings_per_bucket = value; }
    }

    public int NumInfoRecords
    {
        get
        {
            return field_2_sst_info.Count;
        }
    }

    public ExtSSTInfoSubRecord GetInfoRecordAt(int elem)
    {
        return ( ExtSSTInfoSubRecord ) field_2_sst_info[elem];
    }

    public override String ToString()
    {
        StringBuilder buffer = new StringBuilder();

        buffer.Append("[EXTSST]\n");
        buffer.Append("    .dsst           = ")
            .Append(StringUtil.ToHexString(NumStringsPerBucket))
            .Append("\n");
        buffer.Append("    .numInfoRecords = ").Append(NumInfoRecords)
            .Append("\n");
        for (int k = 0; k < NumInfoRecords; k++)
        {
            buffer.Append("    .inforecord     = ").Append(k).Append("\n");
            buffer.Append("    .streampos      = ")
                .Append(StringUtil.ToHexString(
               GetInfoRecordAt(k).StreamPos)).Append("\n");
            buffer.Append("    .sstoffset      = ")
                .Append(StringUtil.ToHexString(
                GetInfoRecordAt(k).BucketSSTOffset))
                    .Append("\n");
        }
        buffer.Append("[/EXTSST]\n");
        return buffer.ToString();
    }

    public override int Serialize(int offset, byte [] data)
    {
        LittleEndian.PutShort(data, 0 + offset, sid);
        LittleEndian.PutShort(data, 2 + offset, (short)(RecordSize - 4));
        LittleEndian.PutShort(data, 4 + offset, field_1_strings_per_bucket);
        int pos = 6;

        for (int k = 0; k < NumInfoRecords; k++)
        {
            ExtSSTInfoSubRecord rec = GetInfoRecordAt(k);
            pos += rec.Serialize(pos + offset, data);
        }
        return pos;
    }

    /** Returns the size of this record */
    public override int RecordSize
    {
        get { return 6 + 8 * NumInfoRecords; }
    }

    public static int GetNumberOfInfoRecsForStrings(int numStrings) {
      int infoRecs = (numStrings / DEFAULT_BUCKET_SIZE);
      if ((numStrings % DEFAULT_BUCKET_SIZE) != 0)
        infoRecs ++;
      //Excel seems to max out after 128 info records.
      //This Isnt really documented anywhere...
      if (infoRecs > MAX_BUCKETS)
        infoRecs = MAX_BUCKETS;
      return infoRecs;
    }

    /** Given a number of strings (in the sst), returns the size of the extsst record*/
    public static int GetRecordSizeForStrings(int numStrings) {
      return 4 + 2 + (GetNumberOfInfoRecsForStrings(numStrings) * 8);
    }

    public override short Sid
    {
        get { return sid; }
    }

    public void SetBucketOffsets( int[] bucketAbsoluteOffsets, int[] bucketRelativeOffsets )
    {
        this.field_2_sst_info = new ArrayList(bucketAbsoluteOffsets.Length);
        for ( int i = 0; i < bucketAbsoluteOffsets.Length; i++ )
        {
            ExtSSTInfoSubRecord r = new ExtSSTInfoSubRecord();
            r.BucketRecordOffset=(short)bucketRelativeOffsets[i];
            r.StreamPos=bucketAbsoluteOffsets[i];
            field_2_sst_info.Add(r);
        }
    }

}
}