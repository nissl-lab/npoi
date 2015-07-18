
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
    using System.Collections.Generic;
    using NPOI.Util;

    /// <summary>
    /// Escher array properties are the most wierd construction ever invented
    /// with all sorts of special cases.  I'm hopeful I've got them all.
    /// @author Glen Stampoultzis (glens at superlinksoftware.com)
    /// </summary>
    public class EscherArrayProperty : EscherComplexProperty, IEnumerable<byte[]> 
    {
        /**
         * The size of the header that goes at the
         *  start of the array, before the data
         */
        private const int FIXED_SIZE = 3 * 2;
        /**
         * Normally, the size recorded in the simple data (for the complex
         *  data) includes the size of the header.
         * There are a few cases when it doesn't though...
         */
        private bool sizeIncludesHeaderSize = true;

        /**
         * When Reading a property from data stream remeber if the complex part is empty and Set this flag.
         */
        private bool emptyComplexPart = false;

        public EscherArrayProperty(short id, byte[] complexData)
            : base(id, CheckComplexData(complexData))
        {

            emptyComplexPart = complexData.Length == 0;
        }

        public EscherArrayProperty(short propertyNumber, bool isBlipId, byte[] complexData)
            : base(propertyNumber, isBlipId, CheckComplexData(complexData))
        {

        }

        private static byte[] CheckComplexData(byte[] complexData)
        {
            if (complexData == null || complexData.Length == 0)
                complexData = new byte[6];

            return complexData;
        }

        public int NumberOfElementsInArray
        {
            get
            {
                if (emptyComplexPart)
                {
                    return 0;
                }
                return LittleEndian.GetUShort(_complexData, 0);
            }
            set
            {
                int expectedArraySize = value * GetActualSizeOfElements(SizeOfElements) + FIXED_SIZE;
                if (expectedArraySize != _complexData.Length)
                {
                    byte[] newArray = new byte[expectedArraySize];
                    Array.Copy(_complexData, 0, newArray, 0, _complexData.Length);
                    _complexData = newArray;
                }
                LittleEndian.PutShort(_complexData, 0, (short)value);
            }
        }
        public int NumberOfElementsInMemory
        {
            get
            {
                return LittleEndian.GetUShort(_complexData, 2);
            }
            set
            {
                int expectedArraySize = value * GetActualSizeOfElements(this.SizeOfElements) + FIXED_SIZE;
                if (expectedArraySize != _complexData.Length)
                {
                    byte[] newArray = new byte[expectedArraySize];
                    Array.Copy(_complexData, 0, newArray, 0, expectedArraySize);
                    _complexData = newArray;
                }
                LittleEndian.PutShort(_complexData, 2, (short)value);
            }
        }

        public short SizeOfElements
        {
            get
            {
                return LittleEndian.GetShort(_complexData, 4);
            }
            set
            {
                LittleEndian.PutShort(_complexData, 4, (short)value);

                int expectedArraySize = NumberOfElementsInArray * GetActualSizeOfElements(SizeOfElements) + FIXED_SIZE;
                if (expectedArraySize != _complexData.Length)
                {
                    // Keep just the first 6 bytes.  The rest is no good to us anyway.
                    byte[] newArray = new byte[expectedArraySize];
                    Array.Copy(_complexData, 0, newArray, 0, 6);
                    _complexData = newArray;
                }
            }
        }

        /// <summary>
        /// Gets the element.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public byte[] GetElement(int index)
        {
            int actualSize = GetActualSizeOfElements(SizeOfElements);
            byte[] result = new byte[actualSize];
            Array.Copy(_complexData, FIXED_SIZE + index * actualSize, result, 0, result.Length);
            return result;
        }

        /// <summary>
        /// Sets the element.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="element">The element.</param>
        public void SetElement(int index, byte[] element)
        {
            int actualSize = GetActualSizeOfElements(SizeOfElements);
            Array.Copy(element, 0, _complexData, FIXED_SIZE + index * actualSize, actualSize);
        }

        /// <summary>
        /// Retrieves the string representation for this property.
        /// </summary>
        /// <returns></returns>
        public override String ToString()
        {
            String nl = Environment.NewLine;

            StringBuilder results = new StringBuilder();
            results.Append("    {EscherArrayProperty:" + nl);
            results.Append("     Num Elements: " + NumberOfElementsInArray + nl);
            results.Append("     Num Elements In Memory: " + NumberOfElementsInMemory + nl);
            results.Append("     Size of elements: " + SizeOfElements + nl);
            for (int i = 0; i < NumberOfElementsInArray; i++)
            {
                results.Append("     Element " + i + ": " + HexDump.ToHex(GetElement(i)) + nl);
            }
            results.Append("}" + nl);

            return "propNum: " + PropertyNumber
                    + ", propName: " + EscherProperties.GetPropertyName(PropertyNumber)
                    + ", complex: " + IsComplex
                    + ", blipId: " + IsBlipId
                    + ", data: " + nl + results.ToString();
        }
        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append("<").Append(GetType().Name).Append(" id=\"0x").Append(HexDump.ToHex(Id))
                    .Append("\" name=\"").Append(Name).Append("\" blipId=\"")
                    .Append(IsBlipId).Append("\">\n");
            for (int i = 0; i < NumberOfElementsInArray; i++)
            {
                builder.Append("\t").Append(tab).Append("<Element>").Append(HexDump.ToHex(GetElement(i))).Append("</Element>\n");
            }
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
        /// <summary>
        /// We have this method because the way in which arrays in escher works
        /// is screwed for seemly arbitary reasons.  While most properties are
        /// fairly consistent and have a predictable array size, escher arrays
        /// have special cases.
        /// </summary>
        /// <param name="data">The data array containing the escher array information</param>
        /// <param name="offset">The offset into the array to start Reading from.</param>
        /// <returns>the number of bytes used by this complex property.</returns>
        public int SetArrayData(byte[] data, int offset)
        {
            if (emptyComplexPart)
            {
                _complexData = new byte[0];
            }
            else
            {
                short numElements = LittleEndian.GetShort(data, offset);
                short numReserved = LittleEndian.GetShort(data, offset + 2);
                short sizeOfElements = LittleEndian.GetShort(data, offset + 4);

                int arraySize = GetActualSizeOfElements(sizeOfElements) * numElements;
                if (arraySize == _complexData.Length)
                {
                    // The stored data size in the simple block excludes the header size
                    _complexData = new byte[arraySize + 6];
                    sizeIncludesHeaderSize = false;
                }
                Array.Copy(data, offset, _complexData, 0, _complexData.Length);
            }
            return _complexData.Length;
        }

        /// <summary>
        /// Serializes the simple part of this property.  ie the first 6 bytes.
        /// Needs special code to handle the case when the size doesn't
        /// include the size of the header block
        /// </summary>
        /// <param name="data"></param>
        /// <param name="pos"></param>
        /// <returns></returns>
        public override int SerializeSimplePart(byte[] data, int pos)
        {
            LittleEndian.PutShort(data, pos, Id);
            int recordSize = _complexData.Length;
            if (!sizeIncludesHeaderSize)
            {
                recordSize -= 6;
            }
            LittleEndian.PutInt(data, pos + 2, recordSize);
            return 6;
        }

        /// <summary>
        /// Sometimes the element size is stored as a negative number.  We
        /// negate it and shift it to Get the real value.
        /// </summary>
        /// <param name="sizeOfElements">The size of elements.</param>
        /// <returns></returns>
        public static int GetActualSizeOfElements(short sizeOfElements)
        {
            if (sizeOfElements < 0)
                return (short)((-sizeOfElements) >> 2);
            else
                return sizeOfElements;
        }


        public IEnumerator<byte[]> GetEnumerator()
        {
            return new EscherArrayEnumerator(this);
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        private class EscherArrayEnumerator : IEnumerator<byte[]>
        {
            EscherArrayProperty dataHolder;
            public EscherArrayEnumerator(EscherArrayProperty eap)
            {
                dataHolder = eap;
            }
            private int idx = -1;
            public byte[] Current
            {
                get
                {
                    if (idx < 0 || idx > dataHolder.NumberOfElementsInArray)
                        throw new IndexOutOfRangeException();
                    return dataHolder.GetElement(idx);
                }
            }

            public void Dispose()
            {
                throw new NotImplementedException();
            }

            object System.Collections.IEnumerator.Current
            {
                get { return this.Current; }
            }

            public bool MoveNext()
            {
                idx++;
                return (idx < dataHolder.NumberOfElementsInArray);
            }

            public void Reset()
            {
                throw new NotImplementedException();
            }
        }
    }
}