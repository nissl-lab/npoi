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

using System;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class Array
    {
        internal sealed class ArrayDimension
        {
            internal long _size;
            public const int SIZE = 8;

            private int _indexOffset;

            public void Read(LittleEndianByteArrayInputStream lei)
            {
                _size = lei.ReadUInt();
                _indexOffset = lei.ReadInt();
            }
        }

        internal sealed class ArrayHeader
        {
            private ArrayDimension[] _dimensions;
            internal int _type;

            public void Read(LittleEndianByteArrayInputStream lei)
            {
                long numDimensionsUnsigned = lei.ReadUInt();

                if(!(1 <= numDimensionsUnsigned && numDimensionsUnsigned <= 31))
                {
                    String msg = "Array dimension number "+numDimensionsUnsigned+" is not in [1; 31] range";
                    throw new IllegalPropertySetDataException(msg);
                }
                int numDimensions = (int) numDimensionsUnsigned;

                _dimensions = new ArrayDimension[numDimensions];
                for(int i = 0; i < numDimensions; i++)
                {
                    ArrayDimension ad = new ArrayDimension();
                    ad.Read(lei);
                    _dimensions[i] = ad;
                }
            }

            public long NumberOfScalarValues
            {
                get
                {
                    long result = 1;
                    foreach(ArrayDimension dimension in _dimensions)
                        result *= dimension._size;
                    return result;
                }
            }

            //public int Size
            //{
            //    get
            //    {
            //        return LittleEndian.INT_SIZE * 2 + _dimensions.Length
            //            * ArrayDimension.SIZE;
            //    }
            //}

            public int Type
            {
                get { return _type; }
            }
        }

        private ArrayHeader _header = new ArrayHeader();
        private TypedPropertyValue[] _values;

        public void Read(LittleEndianByteArrayInputStream lei)
        {
            _header.Read(lei);

            long numberOfScalarsLong = _header.NumberOfScalarValues;
            if(numberOfScalarsLong > int.MaxValue)
            {
                String msg =
                    "Sorry, but POI can't store array of properties with size of " +
                    numberOfScalarsLong + " in memory";
                throw new InvalidOperationException(msg);
            }
            int numberOfScalars = (int) numberOfScalarsLong;

            _values = new TypedPropertyValue[numberOfScalars];
            int paddedType = (_header._type == Variant.VT_VARIANT) ? 0 : _header._type;
            for ( int i = 0; i < numberOfScalars; i++ ) 
            {
                TypedPropertyValue typedPropertyValue = new TypedPropertyValue(paddedType, null);
                typedPropertyValue.Read(lei);
                _values[i] = typedPropertyValue;
                if (paddedType != 0) {
                    TypedPropertyValue.SkipPadding(lei);
                }
            }
        }

        private TypedPropertyValue[] GetValues(){
            return _values;
        }
    }
}