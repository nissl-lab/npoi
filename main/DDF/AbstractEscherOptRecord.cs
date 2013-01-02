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
using System.Collections.Generic;
using NPOI.Util;
using System;
using System.Text;
namespace NPOI.DDF
{

    /**
     * Common abstract class for {@link EscherOptRecord} and
     * {@link EscherTertiaryOptRecord}
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     * @author Glen Stampoultzis
     */
    public abstract class AbstractEscherOptRecord : EscherRecord
    {
        protected List<EscherProperty> properties = new List<EscherProperty>();

        /**
         * Add a property to this record.
         */
        public void AddEscherProperty(EscherProperty prop)
        {
            properties.Add(prop);
        }

        public override int FillFields(byte[] data, int offset,
                IEscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);
            short propertiesCount = ReadInstance(data, offset);
            int pos = offset + 8;

            EscherPropertyFactory f = new EscherPropertyFactory();
            properties = f.CreateProperties(data, pos, propertiesCount);
            return bytesRemaining + 8;
        }

        /**
         * The list of properties stored by this record.
         */
        public List<EscherProperty> EscherProperties
        {
            get
            {
                return properties;
            }
        }

        /**
         * The list of properties stored by this record.
         */
        public EscherProperty GetEscherProperty(int index)
        {
            return properties[index];
        }

        private int PropertiesSize
        {
            get
            {
                int totalSize = 0;
                foreach (EscherProperty property in properties)
                {
                    totalSize += property.PropertySize;
                }

                return totalSize;
            }
        }


        public override int RecordSize
        {
            get
            {
                return 8 + PropertiesSize;
            }
        }

        public EscherProperty Lookup(int propId)
        {
            foreach (EscherProperty prop in properties)
            {
                if (prop.PropertyNumber == propId)
                {

                    return prop;
                }
            }
            return null;
        }

        public override int Serialize(int offset, byte[] data,
                EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            LittleEndian.PutInt(data, offset + 4, PropertiesSize);
            int pos = offset + 8;
            foreach (EscherProperty property in properties)
            {
                pos += property.SerializeSimplePart(data, pos);
            }
            foreach (EscherProperty property in properties)
            {
                pos += property.SerializeComplexPart(data, pos);
            }
            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return pos - offset;
        }
        internal class EscherPropertyComparer : IComparer<EscherProperty>
        {
            public int Compare(EscherProperty p1, EscherProperty p2)
            {
                short s1 = p1.PropertyNumber;
                short s2 = p2.PropertyNumber;
                return s1 < s2 ? -1 : s1 == s2 ? 0 : 1;
            }
        }
        /**
         * Records should be sorted by property number before being stored.
         */
        public void SortProperties()
        {
            properties.Sort(new EscherPropertyComparer());
        }

        /**
         * Set an escher property. If a property with given propId already
         exists it is replaced.
         *
         * @param value the property to set.
         */
        public void SetEscherProperty(EscherProperty value)
        {
            List<EscherProperty> toRemove = new List<EscherProperty>();
            for (IEnumerator<EscherProperty> iterator =
                          properties.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherProperty prop = iterator.Current;
                if (prop.Id == value.Id)
                {
                    //iterator.Remove();
                    toRemove.Add(prop);
                }
            }
            foreach (EscherProperty e in toRemove)
                EscherProperties.Remove(e);
            properties.Add(value);
            SortProperties();
        }

        public void RemoveEscherProperty(int num)
        {
            List<EscherProperty> toRemove = new List<EscherProperty>();
            for (IEnumerator<EscherProperty> iterator = EscherProperties.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherProperty prop = iterator.Current;
                if (prop.PropertyNumber == num)
                {
                    //iterator.Remove();
                    toRemove.Add(prop);
                }
            }
            foreach (EscherProperty e in toRemove)
                EscherProperties.Remove(e);
        }

        /**
         * Retrieve the string representation of this record.
         */
        public override String ToString()
        {
            String nl = Environment.NewLine;

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(GetType().Name);
            stringBuilder.Append(":");
            stringBuilder.Append(nl);
            stringBuilder.Append("  isContainer: ");
            stringBuilder.Append(IsContainerRecord);
            stringBuilder.Append(nl);
            stringBuilder.Append("  version: 0x");
            stringBuilder.Append(HexDump.ToHex(Version));
            stringBuilder.Append(nl);
            stringBuilder.Append("  instance: 0x");
            stringBuilder.Append(HexDump.ToHex(Instance));
            stringBuilder.Append(nl);
            stringBuilder.Append("  recordId: 0x");
            stringBuilder.Append(HexDump.ToHex(RecordId));
            stringBuilder.Append(nl);
            stringBuilder.Append("  numchildren: ");
            stringBuilder.Append(ChildRecords.Count);
            stringBuilder.Append(nl);
            stringBuilder.Append("  properties:");
            stringBuilder.Append(nl);

            foreach (EscherProperty property in properties)
            {
                stringBuilder.Append("    " + property.ToString() + nl);
            }

            return stringBuilder.ToString();
        }

        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name,
                    HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)));
            foreach (EscherProperty property in EscherProperties)
            {
                builder.Append(property.ToXml(tab + "\t"));
            }
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
    }
}

