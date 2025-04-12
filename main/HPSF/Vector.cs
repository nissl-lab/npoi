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
        private readonly short _type;

        private TypedPropertyValue[] _values;

        public Vector(short type)
        {
            this._type = type;
        }

        internal void Read( LittleEndianByteArrayInputStream lei )
        {
            long longLength = lei.ReadUInt();

            if ( longLength > int.MaxValue )
            {
                throw new InvalidOperationException( "Vector is too long -- " + longLength );
            }
            int length = (int) longLength;

            _values = new TypedPropertyValue[length];

            int paddedType = (_type == Variant.VT_VARIANT) ? 0 : _type;
            for ( int i = 0; i < length; i++ )
            {
                TypedPropertyValue value = new TypedPropertyValue(paddedType, null);
                if (paddedType == 0)
                {
                    value.Read(lei);
                }
                else 
                {
                    value.ReadValue(lei);
                }
                _values[i] = value;
            }
        }

        public TypedPropertyValue[] Values
        {
            get { return _values; }
        }
    }
}