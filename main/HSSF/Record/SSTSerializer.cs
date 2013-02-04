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

namespace NPOI.HSSF.Record
{
    using NPOI.HSSF.Record.Cont;
    using NPOI.Util;

    /**
     * This class handles serialization of SST records.  It utilizes the record processor
     * class write individual records. This has been refactored from the SSTRecord class.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class SSTSerializer
    {

        private int _numStrings;
        private int _numUniqueStrings;

        private IntMapper<UnicodeString> strings;

        /** OffSets from the beginning of the SST record (even across continuations) */
        private int[] bucketAbsoluteOffsets;
        /** OffSets relative the start of the current SST or continue record */
        private int[] bucketRelativeOffsets;
        // fix warning CS0169 "never used": int startOfSST, startOfRecord;

        public SSTSerializer(IntMapper<UnicodeString> strings, int numStrings, int numUniqueStrings)
        {
            this.strings = strings;
            _numStrings = numStrings;
            _numUniqueStrings = numUniqueStrings;

            int infoRecs = ExtSSTRecord.GetNumberOfInfoRecsForStrings(strings.Size);
            this.bucketAbsoluteOffsets = new int[infoRecs];
            this.bucketRelativeOffsets = new int[infoRecs];
        }

        public void Serialize(ContinuableRecordOutput out1)
        {
            out1.WriteInt(_numStrings);
            out1.WriteInt(_numUniqueStrings);

            for (int k = 0; k < strings.Size; k++)
            {
                if (k % ExtSSTRecord.DEFAULT_BUCKET_SIZE == 0)
                {
                    int rOff = out1.TotalSize;
                    int index = k / ExtSSTRecord.DEFAULT_BUCKET_SIZE;
                    if (index < ExtSSTRecord.MAX_BUCKETS)
                    {
                        //Excel only indexes the first 128 buckets.
                        bucketAbsoluteOffsets[index] = rOff;
                        bucketRelativeOffsets[index] = rOff;
                    }
                }
                UnicodeString s = GetUnicodeString(k);
                s.Serialize(out1);
            }
        }


        private UnicodeString GetUnicodeString(int index)
        {
            return GetUnicodeString(strings, index);
        }

        private static UnicodeString GetUnicodeString(IntMapper<UnicodeString> strings, int index)
        {
            return (UnicodeString)strings[index];
        }

        public int[] BucketAbsoluteOffsets
        {
            get
            {
                return bucketAbsoluteOffsets;
            }
        }

        public int[] BucketRelativeOffsets
        {
            get
            {
                return bucketRelativeOffsets;
            }
        }
    }
}