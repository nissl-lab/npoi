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
         internal class ArrayDimension
    {
        public const int SIZE = 8;

        private int _indexOffset;
        internal long _size;

        public ArrayDimension( byte[] data, int offset )
        {
            _size = LittleEndian.GetUInt( data, offset );
            _indexOffset = LittleEndian.GetInt( data, offset
                    + LittleEndian.INT_SIZE );
        }
    }

    internal class ArrayHeader
    {
        private ArrayDimension[] _dimensions;
        internal int _type;

        public ArrayHeader( byte[] data, int startOffset )
        {
            int offset = startOffset;

            _type = LittleEndian.GetInt( data, offset );
            offset += LittleEndian.INT_SIZE;

            long numDimensionsUnsigned = LittleEndian.GetUInt( data, offset );
            offset += LittleEndian.INT_SIZE;

            if ( !( 1 <= numDimensionsUnsigned && numDimensionsUnsigned <= 31 ) )
                throw new IllegalPropertySetDataException(
                        "Array dimension number " + numDimensionsUnsigned
                                + " is not in [1; 31] range" );
            int numDimensions = (int) numDimensionsUnsigned;

            _dimensions = new ArrayDimension[numDimensions];
            for ( int i = 0; i < numDimensions; i++ )
            {
                _dimensions[i] = new ArrayDimension( data, offset );
                offset += ArrayDimension.SIZE;
            }
        }

        public long NumberOfScalarValues
        {
            get
            {
                long result = 1;
                foreach (ArrayDimension dimension in _dimensions)
                    result *= dimension._size;
                return result;
            }
        }

        public int Size
        {
            get
            {
                return LittleEndian.INT_SIZE*2 + _dimensions.Length
                       *ArrayDimension.SIZE;
            }
        }

        public int Type
        {
            get { return _type; }
        }
    }

    private ArrayHeader _header;
    private TypedPropertyValue[] _values;

    public Array()
    {
    }

    public Array( byte[] data, int offset )
    {
        Read(data, offset);
    }

    public int Read( byte[] data, int startOffset )
    {
        int offset = startOffset;

        _header = new ArrayHeader( data, offset );
        offset += _header.Size;

        long numberOfScalarsLong = _header.NumberOfScalarValues;
        if ( numberOfScalarsLong > int.MaxValue )
            throw new InvalidOperationException(
                    "Sorry, but POI can't store array of properties with size of "
                            + numberOfScalarsLong + " in memory" );
        int numberOfScalars = (int) numberOfScalarsLong;

        _values = new TypedPropertyValue[numberOfScalars];
        int type = _header._type;
        if ( type == Variant.VT_VARIANT )
        {
            for ( int i = 0; i < numberOfScalars; i++ )
            {
                TypedPropertyValue typedPropertyValue = new TypedPropertyValue();
                offset += typedPropertyValue.Read( data, offset );
            }
        }
        else
        {
            for ( int i = 0; i < numberOfScalars; i++ )
            {
                TypedPropertyValue typedPropertyValue = new TypedPropertyValue(
                        type, null );
                offset += typedPropertyValue.ReadValuePadded( data, offset );
            }
        }

        return offset - startOffset;
    }
    }
}