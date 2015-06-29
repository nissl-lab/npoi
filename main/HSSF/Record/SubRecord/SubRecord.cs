
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
    using System.Text;
    using System.IO;

    using NPOI.Util;

    /**
     * Subrecords are part of the OBJ class.
     */
    public abstract class SubRecord : ICloneable
    {
        public static SubRecord CreateSubRecord(ILittleEndianInput in1, CommonObjectType cmoOt)
        {
            int sid = in1.ReadUShort();
            int secondUShort = in1.ReadUShort(); // Often (but not always) the datasize for the sub-record


            switch (sid)
            {
                case CommonObjectDataSubRecord.sid:
                    return new CommonObjectDataSubRecord(in1, secondUShort);
                case EmbeddedObjectRefSubRecord.sid:
                    return new EmbeddedObjectRefSubRecord(in1, secondUShort);
                case GroupMarkerSubRecord.sid:
                    return new GroupMarkerSubRecord(in1, secondUShort);
                case EndSubRecord.sid:
                    return new EndSubRecord(in1, secondUShort);
                case NoteStructureSubRecord.sid:
                    return new NoteStructureSubRecord(in1, secondUShort);
                case LbsDataSubRecord.sid:
                    return new LbsDataSubRecord(in1, secondUShort, (int)cmoOt);
                case FtCblsSubRecord.sid:
                    return new FtCblsSubRecord(in1, secondUShort);
                case FtPioGrbitSubRecord.sid:
                    return new FtPioGrbitSubRecord(in1, secondUShort);
                case FtCfSubRecord.sid:
                    return new FtCfSubRecord(in1, secondUShort);
            }
            return new UnknownSubRecord(in1, sid, secondUShort);
        }
        public abstract short Sid { get; }
        public abstract int DataSize { get; }
        public abstract void Serialize(ILittleEndianOutput out1);
        public byte[] Serialize()
        {
            int size = DataSize + 4;
            using (MemoryStream baos = new MemoryStream(size))
            {
                Serialize(new LittleEndianOutputStream(baos));
                if (baos.Length != size)
                {
                    throw new Exception("write size mismatch");
                }
                return baos.ToArray();
            }
        }
        /**
 * Wether this record terminates the sub-record stream.
 * There are two cases when this method must be overridden and return <c>true</c>
 *  - EndSubRecord (sid = 0x00)
 *  - LbsDataSubRecord (sid = 0x12)
 *
 * @return whether this record is the last in the sub-record stream
 */
        public virtual bool IsTerminating
        {
            get
            {
                return false;
            }
        }

        public abstract Object Clone();
    }

     public class UnknownSubRecord : SubRecord
     {

         private int _sid;
         private byte[] _data;

         public UnknownSubRecord(ILittleEndianInput in1, int sid, int size)
         {
             _sid = sid;
             byte[] buf = new byte[size];
             in1.ReadFully(buf);
             _data = buf;
         }
         public override int DataSize
         {
             get
             {
                 return _data.Length;
             }
         }
         public override short Sid
         {
             get 
             {
                 return (short)_sid;
             }
         }
         public override void Serialize(ILittleEndianOutput out1)
         {
             out1.WriteShort(_sid);
             out1.WriteShort(_data.Length);
             out1.Write(_data);
         }
         public override Object Clone()
         {
             return this;
         }
         public override String ToString()
         {
             StringBuilder sb = new StringBuilder(64);
             sb.Append(GetType().Name).Append(" [");
             sb.Append("sid=").Append(HexDump.ShortToHex(_sid));
             sb.Append(" size=").Append(_data.Length);
             sb.Append(" : ").Append(HexDump.ToHex(_data));
             sb.Append("]\n");
             return sb.ToString();
         }
     }
}
