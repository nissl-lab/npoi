
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

namespace NPOI.DDF
{
    using System;
    using System.Collections;
    using System.Reflection;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using System.Collections.Generic;

    /// <summary>
    /// Generates escher records when provided the byte array containing those records.
    /// @author Glen Stampoultzis
    /// @author Nick Burch   (nick at torchbox . com)
    /// </summary>
    /// <see cref="EscherRecordFactory"/>
    public class DefaultEscherRecordFactory : EscherRecordFactory
    {
        private static Type[] escherRecordClasses = {
            
            typeof(EscherBSERecord), typeof(EscherOptRecord), typeof(EscherTertiaryOptRecord),
            typeof(EscherClientAnchorRecord), 
            typeof(EscherDgRecord), typeof(EscherSpgrRecord), typeof(EscherSpRecord), 
            typeof(EscherClientDataRecord), typeof(EscherDggRecord),
            typeof(EscherSplitMenuColorsRecord), typeof(EscherChildAnchorRecord), typeof(EscherTextboxRecord)
        };
        private static Dictionary<short,ConstructorInfo> recordsMap = RecordsToMap(escherRecordClasses);

        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultEscherRecordFactory"/> class.
        /// </summary>
        public DefaultEscherRecordFactory()
        {

        }

        /// <summary>
        /// Generates an escher record including the any children contained under that record.
        /// An exception is thrown if the record could not be generated.
        /// </summary>
        /// <param name="data">The byte array containing the records</param>
        /// <param name="offset">The starting offset into the byte array</param>
        /// <returns>The generated escher record</returns>
        public virtual EscherRecord CreateRecord(byte[] data, int offset)
        {
            EscherRecord.EscherRecordHeader header = EscherRecord.EscherRecordHeader.ReadHeader(data, offset);

            // Options of 0x000F means container record
            // However, EscherTextboxRecord are containers of records for the
            //  host application, not of other Escher records, so treat them
            //  differently
            if ((header.Options & (short)0x000F) == (short)0x000F
                 && header.RecordId != EscherTextboxRecord.RECORD_ID)
            {
                EscherContainerRecord r = new EscherContainerRecord();
                r.RecordId = header.RecordId;
                r.Options = header.Options;
                return r;
            }
            if (header.RecordId >= EscherBlipRecord.RECORD_ID_START && header.RecordId <= EscherBlipRecord.RECORD_ID_END)
            {
                EscherBlipRecord r;
                if (header.RecordId == EscherBitmapBlip.RECORD_ID_DIB ||
                        header.RecordId == EscherBitmapBlip.RECORD_ID_JPEG ||
                        header.RecordId == EscherBitmapBlip.RECORD_ID_PNG)
                {
                    r = new EscherBitmapBlip();
                }
                else if (header.RecordId == EscherMetafileBlip.RECORD_ID_EMF ||
                        header.RecordId == EscherMetafileBlip.RECORD_ID_WMF ||
                        header.RecordId == EscherMetafileBlip.RECORD_ID_PICT)
                {
                    r = new EscherMetafileBlip();
                }
                else
                {
                    r = new EscherBlipRecord();
                }
                r.RecordId = header.RecordId;
                r.Options = header.Options;
                return r;
            }

            //ConstructorInfo recordConstructor = (ConstructorInfo) recordsMap[header.RecordId];
            ConstructorInfo recordConstructor = null;
            if (recordsMap.ContainsKey(header.RecordId))
                recordConstructor = recordsMap[header.RecordId];

            EscherRecord escherRecord = null;
            if (recordConstructor == null)
            {
                return new UnknownEscherRecord();
            }

            try
            {
                escherRecord = (EscherRecord)recordConstructor.Invoke(new object[] { });
                //escherRecord = (EscherRecord)Activator.CreateInstance(recordConstructor);
            }
            catch (Exception)
            {
                return new UnknownEscherRecord();
            }
            escherRecord.RecordId = header.RecordId;
            escherRecord.Options = header.Options;
            return escherRecord;

        }

        /// <summary>
        /// Converts from a list of classes into a map that Contains the record id as the key and
        /// the Constructor in the value part of the map.  It does this by using reflection to look up
        /// the RECORD_ID field then using reflection again to find a reference to the constructor.
        /// </summary>
        /// <param name="records">The records to convert</param>
        /// <returns>The map containing the id/constructor pairs.</returns>
        private static Dictionary<short, ConstructorInfo> RecordsToMap(Type[] records)
        {
            Dictionary<short, ConstructorInfo> result = new Dictionary<short, ConstructorInfo>();
            //ConstructorInfo constructor;
            Type[] EMPTY_CLASS_ARRAY = new Type[0];
            for (int i = 0; i < records.Length; i++)
            {
                Type recordType = records[i];
                short sid = 0;

                try
                {
                    sid = (short)recordType.GetField("RECORD_ID").GetValue(null);
                }
                catch (Exception)
                {
                    throw new RecordFormatException(
                            "Unable to determine record types");
                }
                ConstructorInfo ci;
                try
                {
                    ci = recordType.GetConstructor(EMPTY_CLASS_ARRAY);
                }
                catch(Exception e)
                {
                    throw new RuntimeException(e);
                }
                //result[sid] = recordType;        //constructor;
                result.Add(sid, ci);
            }
            return result;
        }

    }
}