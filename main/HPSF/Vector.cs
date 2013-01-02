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
    public class Vector
    {
        private short _type;

        private TypedPropertyValue[] _values;

        public Vector(byte[] data, int startOffset, short type)
        {
            this._type = type;
            Read(data, startOffset);
        }

        public Vector(short type)
        {
            this._type = type;
        }

        public int Read(byte[] data, int startOffset)
        {
            int offset = startOffset;

            long longLength = LittleEndian.GetUInt(data, offset);
            offset += LittleEndian.INT_SIZE;

            if (longLength > int.MaxValue)
                throw new InvalidOperationException("Vector is too long -- "
                        + longLength);
            int length = (int)longLength;

            _values = new TypedPropertyValue[length];

            if (_type == Variant.VT_VARIANT)
            {
                for (int i = 0; i < length; i++)
                {
                    TypedPropertyValue value = new TypedPropertyValue();
                    offset += value.Read(data, offset);
                    _values[i] = value;
                }
            }
            else
            {
                for (int i = 0; i < length; i++)
                {
                    TypedPropertyValue value = new TypedPropertyValue(_type, null);
                    // be aware: not padded here
                    offset += value.ReadValue(data, offset);
                    _values[i] = value;
                }
            }
            return offset - startOffset;
        }

        public TypedPropertyValue[] Values
        {
            get { return _values; }
        }
    }
}