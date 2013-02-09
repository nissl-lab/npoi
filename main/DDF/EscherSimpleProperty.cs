
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

using System.Text;

namespace NPOI.DDF
{
    using System;
    using NPOI.Util;

    /// <summary>
    /// A simple property is of fixed Length and as a property number in Addition
    /// to a 32-bit value.  Properties that can't be stored in only 32-bits are
    /// stored as EscherComplexProperty objects.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class EscherSimpleProperty : EscherProperty
    {
        protected int propertyValue;

        /// <summary>
        /// The id is distinct from the actual property number.  The id includes the property number the blip id
        /// flag and an indicator whether the property is complex or not.
        /// </summary>
        /// <param name="id">The id.</param>
        /// <param name="propertyValue">The property value.</param>
        public EscherSimpleProperty(short id, int propertyValue):base(id)
        {
            this.propertyValue = propertyValue;
        }

        /// <summary>
        /// Constructs a new escher property.  The three parameters are combined to form a property
        /// id.
        /// </summary>
        /// <param name="propertyNumber">The property number.</param>
        /// <param name="isComplex">if set to <c>true</c> [is complex].</param>
        /// <param name="isBlipId">if set to <c>true</c> [is blip id].</param>
        /// <param name="propertyValue">The property value.</param>
        public EscherSimpleProperty(short propertyNumber, bool isComplex, bool isBlipId, int propertyValue):base(propertyNumber, isComplex, isBlipId)
        {
            
            this.propertyValue = propertyValue;
        }

        /// <summary>
        /// Serialize the simple part of the escher record.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="offset">The off set.</param>
        /// <returns>the number of bytes Serialized.</returns>
        public override int SerializeSimplePart(byte[] data, int offset)
        {
            LittleEndian.PutShort(data, offset, Id);
            LittleEndian.PutInt(data, offset + 2, propertyValue);
            return 6;
        }

        /// <summary>
        /// Escher properties consist of a simple fixed Length part and a complex variable Length part.
        /// The fixed Length part is Serialized first.
        /// </summary>
        /// <param name="data"></param>
        /// <param name="pos"></param>
        /// <returns></returns>
        public override int SerializeComplexPart(byte[] data, int pos)
        {
            return 0;
        }

        /// <summary>
        /// Return the 32 bit value of this property.
        /// </summary>
        /// <value>The property value.</value>
        public int PropertyValue
        {
            get { return propertyValue; }
            internal set { propertyValue = value; }
        }

        /// <summary>
        /// Returns true if one escher property is equal to another.
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns></returns>
        public override bool Equals(Object o)
        {
            if (this == o) return true;
            if (!(o is EscherSimpleProperty)) return false;

            EscherSimpleProperty escherSimpleProperty = (EscherSimpleProperty)o;

            if (propertyValue != escherSimpleProperty.propertyValue) return false;
            if (Id != escherSimpleProperty.Id) return false;

            return true;
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            return propertyValue;
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            return "propNum: " + PropertyNumber
                    + ", RAW: 0x" + HexDump.ToHex(Id)
                    + ", propName: " + EscherProperties.GetPropertyName(PropertyNumber)
                    + ", complex: " + IsComplex
                    + ", blipId: " + IsBlipId
                    + ", value: " + propertyValue + " (0x" + HexDump.ToHex(propertyValue) + ")";
        }
        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append("<").Append(GetType().Name).Append(" id=\"0x").Append(HexDump.ToHex(Id))
                    .Append("\" name=\"").Append(Name).Append("\" blipId=\"")
                    .Append(IsBlipId).Append("\" complex=\"").Append(IsComplex).Append("\" value=\"").Append("0x")
                    .Append(HexDump.ToHex(propertyValue)).Append("\"/>\n");
            return builder.ToString();
        }
    }
}