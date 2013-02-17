
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

    using System;
    using System.Collections;
    using System.Text;
    using NPOI.Util;
    using NPOI.HSSF.Record.Cont;


    /**
     * Title:        Static String Table Record
     * 
     * Description:  This holds all the strings for LabelSSTRecords.
     * 
     * REFERENCE:    PG 389 Microsoft Excel 97 Developer's Kit (ISBN:
     *               1-57231-498-2)
     * 
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     *
     * @see org.apache.poi.hssf.record.LabelSSTRecord
     * @see org.apache.poi.hssf.record.ContinueRecord
     */

    public class SSTRecord : ContinuableRecord
    {
        public const short sid = 0x00FC;
        private static readonly UnicodeString EMPTY_STRING = new UnicodeString("");

        /** how big can an SST record be? As big as any record can be: 8228 bytes */
        public const int MAX_RECORD_SIZE = 8228;

        /** standard record overhead: two shorts (record id plus data space size)*/
        public const int STD_RECORD_OVERHEAD =  2 * LittleEndianConsts.SHORT_SIZE;

        /** SST overhead: the standard record overhead, plus the number of strings and the number of Unique strings -- two ints */
        public const int SST_RECORD_OVERHEAD =                (STD_RECORD_OVERHEAD + (2 * LittleEndianConsts.INT_SIZE));

        /** how much data can we stuff into an SST record? That would be _max minus the standard SST record overhead */
        public const int MAX_DATA_SPACE = RecordInputStream.MAX_RECORD_DATA_SIZE - 8;//MAX_RECORD_SIZE - SST_RECORD_OVERHEAD;

        // /** overhead for each string includes the string's Char count (a short) and the flag describing its Charistics (a byte) */
        //public const int STRING_MINIMAL_OVERHEAD = LittleEndianConsts.SHORT_SIZE + LittleEndianConsts.BYTE_SIZE;

        /** Union of strings in the SST and EXTSST */
        private int field_1_num_strings;

        /** according to docs ONLY SST */
        private int field_2_num_unique_strings;
        private IntMapper<UnicodeString> field_3_strings;

        private SSTDeserializer deserializer;

        /** Offsets from the beginning of the SST record (even across continuations) */
        int[] bucketAbsoluteOffsets;
        /** Offsets relative the start of the current SST or continue record */
        int[] bucketRelativeOffsets;

        /**
         * default constructor
         */
        public SSTRecord()
        {
            field_1_num_strings = 0;
            field_2_num_unique_strings = 0;
            field_3_strings = new IntMapper<UnicodeString>();
            deserializer = new SSTDeserializer(field_3_strings);
        }

        /**
         * Constructs an SST record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public SSTRecord(RecordInputStream in1)
        {
            // this method Is ALWAYS called after construction -- using
            // the nontrivial constructor, of course -- so this Is where
            // we initialize our fields
            field_1_num_strings = in1.ReadInt();
            field_2_num_unique_strings = in1.ReadInt();
            field_3_strings = new IntMapper<UnicodeString>();
            deserializer = new SSTDeserializer(field_3_strings);
            deserializer.ManufactureStrings(field_2_num_unique_strings, in1);
        }

        /**
         * Add a string.
         *
         * @param string string to be Added
         *
         * @return the index of that string in the table
         */

        public int AddString(UnicodeString str)
        {
            field_1_num_strings++;
            UnicodeString ucs = (str == null) ? EMPTY_STRING
                    : str;
            int rval;
            int index = field_3_strings.GetIndex(ucs);

            if (index != -1)
            {
                rval = index;
            }
            else
            {
                // This is a new string -- we didn't see it among the
                // strings we've already collected
                rval = field_3_strings.Size;
                field_2_num_unique_strings++;
                SSTDeserializer.AddToStringTable(field_3_strings, ucs);
            }
            return rval;
        }

        /**
         * @return number of strings
         */

        public int NumStrings
        {
            get { return field_1_num_strings; }
            set { field_1_num_strings = value; }
        }

        /**
         * @return number of Unique strings
         */

        public int NumUniqueStrings
        {
            get { return field_2_num_unique_strings; }
            set { field_2_num_unique_strings = value; }
        }

        /**
         * Get a particular string by its index
         *
         * @param id index into the array of strings
         *
         * @return the desired string
         */

        public UnicodeString GetString(int id)
        {
            return (UnicodeString)field_3_strings[id];
        }

        public bool IsString16bit(int id)
        {
            UnicodeString unicodeString = ((UnicodeString)field_3_strings[id]);
            return ((unicodeString.OptionFlags & 0x01) == 1);
        }

        /**
         * Return a debugging string representation
         *
         * @return string representation
         */

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SST]\n");
            buffer.Append("    .numstrings     = ")
                    .Append(StringUtil.ToHexString(NumStrings)).Append("\n");
            buffer.Append("    .uniquestrings  = ")
                    .Append(StringUtil.ToHexString(NumUniqueStrings)).Append("\n");
            for (int k = 0; k < field_3_strings.Size; k++)
            {
                UnicodeString s = (UnicodeString)field_3_strings[k];
                buffer.Append("    .string_" + k + "      = ")
                        .Append(s.GetDebugInfo()).Append("\n");
            }
            buffer.Append("[/SST]\n");
            return buffer.ToString();
        }

        /**
         * @return sid
         */
        public override short Sid
        {
            get { return sid; }
        }

        /**
         * @return hashcode
         */
        public override int GetHashCode()
    {
        return field_2_num_unique_strings;
    }

        public override bool Equals(Object o)
        {
            if ((o == null) || (o.GetType() != this.GetType()))
            {
                return false;
            }
            SSTRecord other = (SSTRecord)o;

            return ((field_1_num_strings == other
                    .field_1_num_strings) && (field_2_num_unique_strings == other
                    .field_2_num_unique_strings) && field_3_strings
                    .Equals(other.field_3_strings));
        }

        /**
         * @return an iterator of the strings we hold. All instances are
         *         UnicodeStrings
         */

        public IEnumerator GetStrings()
        {
            return field_3_strings.GetEnumerator(); 
        }

        /**
         * @return count of the strings we hold.
         */

        public int CountStrings
        {
            get { return field_3_strings.Size; }
        }

        /**
         * called by the class that Is responsible for writing this sucker.
         * Subclasses should implement this so that their data Is passed back in a
         * byte array.
         *
         * @return size
         */

    protected override void Serialize(ContinuableRecordOutput out1) {
        SSTSerializer serializer = new SSTSerializer(field_3_strings, NumStrings, NumUniqueStrings );
        serializer.Serialize(out1);
        bucketAbsoluteOffsets = serializer.BucketAbsoluteOffsets;
        bucketRelativeOffsets = serializer.BucketRelativeOffsets;
    }

        SSTDeserializer GetDeserializer()
        {
            return deserializer;
        }

        /**
         * Creates an extended string record based on the current contents of
         * the current SST record.  The offset within the stream to the SST record
         * Is required because the extended string record points directly to the
         * strings in the SST record.
         * 
         * NOTE: THIS FUNCTION MUST ONLY BE CALLED AFTER THE SST RECORD HAS BEEN
         *       SERIALIZED.
         *
         * @param sstOffset     The offset in the stream to the start of the
         *                      SST record.
         * @return  The new SST record.
         */
        public ExtSSTRecord CreateExtSSTRecord(int sstOffset)
        {
            if (bucketAbsoluteOffsets == null || bucketAbsoluteOffsets == null)
                throw new InvalidOperationException("SST record has not yet been Serialized.");

            ExtSSTRecord extSST = new ExtSSTRecord();
            extSST.NumStringsPerBucket=((short)8);
            int[] absoluteOffsets = (int[])bucketAbsoluteOffsets.Clone();
            int[] relativeOffsets = (int[])bucketRelativeOffsets.Clone();
            for (int i = 0; i < absoluteOffsets.Length; i++)
                absoluteOffsets[i] += sstOffset;
            extSST.SetBucketOffsets(absoluteOffsets, relativeOffsets);
            return extSST;
        }

        /**
         * Calculates the size in bytes of the EXTSST record as it would be if the
         * record was Serialized.
         *
         * @return  The size of the ExtSST record in bytes.
         */
        public int CalcExtSSTRecordSize()
        {
            return ExtSSTRecord.GetRecordSizeForStrings(field_3_strings.Size);
        }
    }
}