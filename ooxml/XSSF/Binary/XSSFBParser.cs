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


using System.Collections;
using System.IO;

namespace NPOI.XSSF.Binary
{
    using NPOI.Util;

    /// <summary>
    /// Experimental parser for Microsoft's ooxml xlsb format.
    /// Not thread safe, obviously.  Need to create a new one
    /// for each thread.
    /// </summary>
    /// @since 3.16-beta3
    public abstract class XSSFBParser
    {
        private  LittleEndianInputStream is1;
        private  BitArray records;

        protected XSSFBParser(Stream is1)
        {
            this.is1 = new LittleEndianInputStream(is1);
            records = null;
        }

        /// <summary>
        /// </summary>
        /// <param name="is1">inputStream</param>
        /// <param name="bitSet">call <see cref="HandleRecord(int, byte[])" /> only on those records in this bitSet</param>
        protected XSSFBParser(Stream is1, BitArray bitSet)
        {
            this.is1 = new LittleEndianInputStream(is1);
            records = bitSet;
        }

        public void Parse()
        {
            while(true)
            {
                int bInt = is1.Read();
                if(bInt == -1)
                {
                    return;
                }
                ReadNext((sbyte) bInt);
            }
        }

        private void ReadNext(sbyte b1)
        {

            int recordId = 0;

            //if highest bit == 1
            if((b1 >> 7 & 1) == 1)
            {
                sbyte b2 = (sbyte)is1.ReadByte();
                b1 &= unchecked((sbyte)~(1<<7)); //unset highest bit
                b2 &= unchecked((sbyte)~(1<<7)); //unset highest bit (if it exists?)
                recordId = ((int) b2 << 7)+(int) b1;
            }
            else
            {
                recordId = (int) b1;
            }

            long recordLength = 0;
            int i = 0;
            bool halt = false;
            while(i < 4 && !halt)
            {
                sbyte b = (sbyte)is1.ReadByte();
                halt = (b >> 7 & 1) == 0; //if highest bit !=1 then continue
                b &= unchecked((sbyte)~(1<<7));
                recordLength += (int) b << (i*7); //multiply by 128^i
                i++;

            }
            if(records == null || records.Get(recordId))
            {
                //add sanity check for length?
                byte[] buff = new byte[(int) recordLength];
                is1.ReadFully(buff);
                HandleRecord(recordId, buff);
            }
            else
            {
                long length = is1.Skip(recordLength);
                if(length != recordLength)
                {
                    throw new XSSFBParseException("End of file reached before expected.\t"+
                    "Tried to skip "+recordLength + ", but only skipped "+length);
                }
            }
        }

        //It hurts, hurts, hurts to create a new byte array for every record.
        //However, on a large Excel spreadsheet, this parser was 1/3 faster than
        //the ooxml sax parser (5 seconds for xssfb and 7.5 seconds for xssf.
        //The code is far cleaner to have the parser read all
        //of the data rather than having every component promise that it will read
        //the correct amount.
        public abstract void HandleRecord(int recordType, byte[] data);


    }
}

