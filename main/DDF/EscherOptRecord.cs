
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
    using System.Text;
    using System.Collections;
    using System.Collections.Generic;
    using NPOI.Util;


    /// <summary>
    /// The opt record is used to store property values for a shape.  It is the key to determining
    /// the attributes of a shape.  Properties can be of two types: simple or complex.  Simple types
    /// are fixed Length.  Complex properties are variable Length.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherOptRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF00B);
        public const String RECORD_DESCRIPTION = "msofbtOPT";

        private List<EscherProperty> properties = new List<EscherProperty>();

        /// <summary>
        /// This method deSerializes the record from a byte array.
        /// </summary>
        /// <param name="data">The byte array containing the escher record information</param>
        /// <param name="offset">The starting offset into data</param>
        /// <param name="recordFactory">May be null since this is not a container record.</param>
        /// <returns>The number of bytes Read from the byte array.</returns>
        public override int FillFields(byte[] data, int offset, EscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);
            int pos = offset + 8;

            EscherPropertyFactory f = new EscherPropertyFactory();
            properties = f.CreateProperties(data, pos, GetInstance());
            return bytesRemaining + 8;
        }

        /// <summary>
        /// This method Serializes this escher record into a byte array
        /// </summary>
        /// <param name="offset">The offset into data
        ///  to start writing the record data to.</param>
        /// <param name="data">The byte array to Serialize to.</param>
        /// <returns>The number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            LittleEndian.PutInt(data, offset + 4, PropertiesSize);
            int pos = offset + 8;
            for (IEnumerator iterator = properties.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherProperty escherProperty = (EscherProperty)iterator.Current;
                pos += escherProperty.SerializeSimplePart(data, pos);
            }
            for (IEnumerator iterator = properties.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherProperty escherProperty = (EscherProperty)iterator.Current;
                pos += escherProperty.SerializeComplexPart(data, pos);
            }
            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return pos - offset;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + PropertiesSize; }
        }

        /// <summary>
        /// Automatically recalculate the correct option
        /// </summary>
        /// <value></value>
        public override short Options
        {
            get
            {
                base.Options = (short)((properties.Count<< 4) | 0x3);
                return base.Options;
            }
        }

        /// <summary>
        /// The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "Opt"; }
        }

        /// <summary>
        /// Gets the size of the properties.
        /// </summary>
        /// <returns></returns>
        private int PropertiesSize
        {
            get{
                int totalSize = 0;
                for (IEnumerator iterator = properties.GetEnumerator(); iterator.MoveNext(); )
                {
                    EscherProperty escherProperty = (EscherProperty)iterator.Current;
                    totalSize += escherProperty.PropertySize;
                }
                return totalSize;
            }
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            String nl = Environment.NewLine;
            StringBuilder propertiesBuf = new StringBuilder();
            for (IEnumerator iterator = properties.GetEnumerator(); iterator.MoveNext(); )
                propertiesBuf.Append("    "
                        + iterator.Current.ToString()
                        + nl);

            return "EscherOptRecord:" + nl +
                    "  isContainer: " + IsContainerRecord + nl +
                    "  options: 0x" + HexDump.ToHex(Options) + nl +
                    "  recordId: 0x" + HexDump.ToHex(RecordId) + nl +
                    "  numchildren: " + ChildRecords.Count + nl +
                    "  properties:" + nl +
                    propertiesBuf.ToString();
        }

        /// <summary>
        /// The list of properties stored by this record.
        /// </summary>
        /// <returns></returns>
        public List<EscherProperty> EscherProperties
        {
            get{return properties;}
        }

        /// <summary>
        /// The list of properties stored by this record.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public EscherProperty GetEscherProperty(int index)
        {
            return (EscherProperty)properties[index];
        }

        /// <summary>
        /// Add a property to this record.
        /// </summary>
        /// <param name="prop">The prop.</param>
        public void AddEscherProperty(EscherProperty prop)
        {
            properties.Add(prop);
        }

        public EscherProperty Lookup(int propId)
        {
            foreach (EscherProperty prop in properties)
            {
                if (prop.PropertyNumber == propId)
                    return prop;
            }
            return null;
        }


        int CompareEscherProperty(EscherProperty o1, EscherProperty o2)
        {
            EscherProperty p1 = o1;
            EscherProperty p2 = o2;
            return p1.PropertyNumber.CompareTo(p2.PropertyNumber);
        }

        /// <summary>
        /// Records should be sorted by property number before being stored.
        /// </summary>
        public void SortProperties()
        {
            properties.Sort(CompareEscherProperty);
        }
    }
}
