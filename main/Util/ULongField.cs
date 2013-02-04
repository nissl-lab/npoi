
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.IO;

namespace NPOI.Util
{
    /// <summary>
    /// 
    /// </summary>
    [Obsolete]
    public class ULongField
    {
        private ulong      _value;
        private int _offset;

        /// <summary>
        /// construct the <see cref="LongField"/> with its offset into its containing byte array
        /// </summary>
        /// <param name="offset">The offset.</param>
        public ULongField(int offset)
        {
            if (offset < 0)
            {
                throw new IndexOutOfRangeException("Illegal offset: " + offset);
            }
            _offset = offset;
        }

        /// <summary>
        /// construct the LongField with its offset into its containing
        /// byte array and initialize its value
        /// </summary>
        /// <param name="offset">offset of the field within its byte array</param>
        /// <param name="value">the initial value</param>
        public ULongField(int offset, ulong value):this(offset)
        {
            this.Value=value;
        }

        /// <summary>
        /// Construct the <see cref="LongField"/> class with its offset into its containing
        /// byte array and initialize its value from its byte array
        /// </summary>
        /// <param name="offset">The offset of the field within its byte array</param>
        /// <param name="data">the byte array to read the value from</param>
        public ULongField(int offset, byte [] data)
        {
            this._offset=offset;
            ReadFromBytes(data);
        }

        /// <summary>
        /// construct the <see cref="LongField"/> class with its offset into its containing
        /// byte array, initialize its value, and write the value to a byte
        /// array
        /// </summary>
        /// <param name="offset">offset of the field within its byte array</param>
        /// <param name="value">the initial value</param>
        /// <param name="data">the byte array to write the value to</param>
        public ULongField(int offset, ulong value, byte [] data)
        {
            this._offset=offset;
            Set(value, data);
        }

        /// <summary>
        /// Getg or sets the LongField's current value
        /// </summary>
        /// <value>The current value</value>
        public ulong Value
        {
            get{return _value;}
            set{_value = value;}
        }

        /// <summary>
        /// set the LongField's current value and write it to a byte array
        /// </summary>
        /// <param name="value">value to be set</param>
        /// <param name="data">the byte array to write the value to</param>
        public void Set(ulong value, byte [] data)
        {
            this._value = value;
            WriteToBytes(data);
        }

        /// <summary>
        /// set the value from its offset into an array of bytes
        /// </summary>
        /// <param name="data">the byte array from which the value is to be read</param>
        public void ReadFromBytes(byte [] data)
        {
            _value = LittleEndian.GetULong(data, _offset);
        }

        /// <summary>
        /// set the value from an Stream
        /// </summary>
        /// <param name="stream">the Stream from which the value is to be</param>
        public void ReadFromStream(Stream stream)
        {
            _value = LittleEndian.ReadULong(stream);
        }

        /// <summary>
        /// write the value out to an array of bytes at the appropriate offset
        /// </summary>
        /// <param name="data">the array of bytes to which the value is to be written</param>
        public void WriteToBytes(byte [] data)
        {
            LittleEndian.PutULong(data, _offset, _value);
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            return Convert.ToString(_value);
        }
    }
}
