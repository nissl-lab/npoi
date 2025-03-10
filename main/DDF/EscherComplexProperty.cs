
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
    using System.IO;
    using NPOI.Util;

    /// <summary>
    /// A complex property differs from a simple property in that the data can not fit inside a 32 bit
    /// integer.  See the specification for more detailed information regarding exactly what is
    /// stored here.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherComplexProperty : EscherProperty
    {
        private byte[] _complexData = Array.Empty<byte>();

        /// <summary>
        /// Create a complex property using the property id and a byte array containing the complex
        /// data value.
        /// </summary>
        /// <param name="id"> The id consists of the property number, a flag indicating whether this is a blip id and a flag
        /// indicating that this is a complex property.</param>
        /// <param name="complexData">The value of this property.</param>
        public EscherComplexProperty(short id, byte[] complexData)
            : base(id)
        {
            if (complexData == null)
            {
                throw new ArgumentNullException("complexData can't be null");
            }
            this._complexData = (byte[])complexData.Clone();
        }

        /// <summary>
        /// Create a complex property using the property number, a flag to indicate whether this is a
        /// blip reference and the complex property data.
        /// </summary>
        /// <param name="propertyNumber">The property number.</param>
        /// <param name="isBlipId">Whether this is a blip id.  Should be false.</param>
        /// <param name="complexData">The value of this complex property.</param> 
        public EscherComplexProperty(short propertyNumber, bool isBlipId, byte[] complexData)
            : base(propertyNumber, true, isBlipId)
        {
            if (complexData == null)
            {
                throw new ArgumentNullException("complexData can't be null");
            }
            this._complexData = (byte[])complexData.Clone();
        }

        /// <summary>
        /// Serializes the simple part of this property.  ie the first 6 bytes.
        /// </summary>
        /// <param name="data"></param>
        /// <param name="pos"></param>
        /// <returns></returns>
        public override int SerializeSimplePart(byte[] data, int pos)
        {
            LittleEndian.PutShort(data, pos, Id);
            LittleEndian.PutInt(data, pos + 2, _complexData.Length);
            return 6;
        }

        /// <summary>
        /// Serializes the complex part of this property
        /// </summary>
        /// <param name="data">The data array to Serialize to</param>
        /// <param name="pos">The offset within data to start serializing to.</param>
        /// <returns>The number of bytes Serialized.</returns>
        public override int SerializeComplexPart(byte[] data, int pos)
        {
            Array.Copy(_complexData, 0, data, pos, _complexData.Length);
            return _complexData.Length;
        }

        /// <summary>
        /// Gets the complex data.
        /// </summary>
        /// <value>The complex data.</value>
        public byte[] ComplexData
        {
            get { return _complexData; }
        }

        protected void ResizeComplexData(int newSize, int bytesToKeep)
        {
            if (newSize == _complexData.Length)
            {
                return;
            }
            byte[] newArray = new byte[newSize];
            Array.Copy(_complexData, 0, newArray, 0, Math.Min(bytesToKeep, newSize));
            _complexData = newArray;
        }

        /// <summary>
        /// Determine whether this property is equal to another property.
        /// </summary>
        /// <param name="o">The object to compare to.</param>
        /// <returns>True if the objects are equal.</returns>
        public override bool Equals(Object o)
        {
            if (this == o) return true;
            if (o == null || o is not EscherComplexProperty escherComplexProperty) return false;

            if (!Arrays.Equals(_complexData, escherComplexProperty._complexData)) return false;

            return true;
        }

        /// <summary>
        /// Caclulates the number of bytes required to Serialize this property.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int PropertySize
        {
            get { return 6 + _complexData.Length; }
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            return Id * 11;
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            String dataStr;
            using (MemoryStream b = new MemoryStream())
            {
                try
                {
                    HexDump.Dump(this._complexData, 0, b, 0);
                    dataStr = b.ToString();
                }
                catch (Exception e)
                {
                    dataStr = e.ToString();
                }
            }
            return "propNum: " + PropertyNumber
                    + ", propName: " + EscherProperties.GetPropertyName(PropertyNumber)
                    + ", complex: " + IsComplex
                    + ", blipId: " + IsBlipId
                    + ", data: " + Environment.NewLine + dataStr;
        }
        public override String ToXml(String tab)
        {
            String dataStr = HexDump.ToHex(_complexData, 32);
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append("<").Append(GetType().Name).Append(" id=\"0x").Append(HexDump.ToHex(Id))
                    .Append("\" name=\"").Append(Name).Append("\" blipId=\"")
                    .Append(IsBlipId).Append("\">\n");
            //builder.Append("\t").Append(tab).Append(dataStr);
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
    }
}