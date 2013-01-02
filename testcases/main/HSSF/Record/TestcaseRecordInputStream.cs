/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.HSSF.Record
{
    using System;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using System.Text;


    /**
     * A Record Input Stream derivative that makes access to byte arrays used in the
     * test cases work a bit easier.
     * Creates the sream and moves to the first record.
     *
     * @author Jason Height (jheight at apache.org)
     */
    public class TestcaseRecordInputStream {
    	
	    private TestcaseRecordInputStream() {
		    // no instances of this class
	    }
        /// <summary>
        /// Prepends a mock record identifier to the supplied data and opens a record input stream
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static ILittleEndianInput CreateLittleEndian(byte[] data)
        {
            return new LittleEndianByteArrayInputStream(data);

        }
	    public static RecordInputStream Create(int sid, byte[] data) {
		    return Create(MergeDataAndSid(sid, data.Length, data));
	    }
        /// <summary>
        ///First 4 bytes of data are assumed to be record identifier and length. The supplied 
	    ///data can contain multiple records (sequentially encoded in the same way) 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
	    public static RecordInputStream Create(byte[] data) {
		    Stream inputStream = new MemoryStream(data);
		    RecordInputStream result = new RecordInputStream(inputStream);
		    result.NextRecord();
		    return result;
	    }


        public static byte[] MergeDataAndSid(int sid, int length, byte[] data)
        {
          byte[] result = new byte[data.Length + 4];
          LittleEndian.PutUShort(result, 0, sid);
          LittleEndian.PutUShort(result, 2, length);
          Array.Copy(data, 0, result, 4, data.Length);
          return result;
        }
        /// <summary>
        /// Confirms data sections are equal
        /// </summary>
        /// <param name="msgPrefix">message prefix to be displayed in case of failure</param>
        /// <param name="expectedSid"></param>
        /// <param name="expectedData">just raw data (without ushort sid, ushort size)</param>
        /// <param name="actualRecordBytes">this includes 4 prefix bytes (sid & size)</param>
        public static void ConfirmRecordEncoding(String msgPrefix, int expectedSid, byte[] expectedData, byte[] actualRecordBytes)
        {
            int expectedDataSize = expectedData.Length;
            Assert.AreEqual(actualRecordBytes.Length - 4, expectedDataSize,"Size of encode data mismatch");
            Assert.AreEqual(expectedSid, LittleEndian.GetShort(actualRecordBytes, 0));
            Assert.AreEqual(expectedDataSize, LittleEndian.GetShort(actualRecordBytes, 2));
            for (int i = 0; i < expectedDataSize; i++)
                if (expectedData[i] != actualRecordBytes[i + 4])
                {
                    StringBuilder sb = new StringBuilder(64);
                    if (msgPrefix != null)
                    {
                        sb.Append(msgPrefix).Append(": ");
                    }
                    sb.Append("At offset ").Append(i);
                    sb.Append(": expected ").Append(HexDump.ByteToHex(expectedData[i]));
                    sb.Append(" but found ").Append(HexDump.ByteToHex(actualRecordBytes[i + 4]));
                    throw new AssertionException(sb.ToString());
                }
        }
        /// <summary>
        /// Confirms data sections are equal
        /// </summary>
        /// <param name="expectedSid">just raw data (without sid or size short ints)</param>
        /// <param name="expectedData"></param>
        /// <param name="actualRecordBytes">this includes 4 prefix bytes (sid & size)</param>
        public static void ConfirmRecordEncoding(int expectedSid, byte[] expectedData, byte[] actualRecordBytes)
        {
            ConfirmRecordEncoding(null, expectedSid, expectedData, actualRecordBytes);
        }
    }
}