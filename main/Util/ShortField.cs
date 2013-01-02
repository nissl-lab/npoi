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
using System.Globalization;

namespace NPOI.Util
{
    public class ShortField
    {
        private short     _value;
        private int _offset;

        /// <summary>
        /// construct the ShortField with its offset into its containing
        /// byte array
        /// </summary>
        /// <param name="offset">offset of the field within its byte array</param>
        ///<exception cref="IndexOutOfRangeException">if offset is negative</exception>
        public ShortField(int offset)
        {
            if (offset < 0)
            {
                throw new IndexOutOfRangeException("Illegal offset: "
                                                         + offset);
            }
            _offset = offset;
        }

        /// <summary>
        /// construct the ShortField with its offset into its containing byte array and initialize its value
        /// </summary>
        /// <param name="offset">offset of the field within its byte array</param>
        /// <param name="value">the initial value</param>
        /// <exception cref="IndexOutOfRangeException">if offset is negative</exception> 
        public ShortField(int offset, short value):this(offset)
        {
            this._value = value;
        }

        /// <summary>
        /// Construct the ShortField with its offset into its containing
        /// byte array and initialize its value from its byte array
        /// </summary>
        /// <param name="offset">offset of the field within its byte array</param>
        /// <param name="data">the byte array to read the value from</param>
        /// <exception cref="IndexOutOfRangeException">if the offset is not
        /// within the range of 0..(data.length - 1)</exception> 
        public ShortField(int offset, byte[] data)
            : this(offset)
        {
            ReadFromBytes(data);
        }

        /// <summary>
        /// construct the ShortField with its offset into its containing
        /// byte array, initialize its value, and write its value to its
        /// byte array
        /// </summary>
        /// <param name="offset">offset of the field within its byte array</param>
        /// <param name="value">the initial value</param>
        /// <param name="data">the byte array to write the value to</param>
        /// <exception cref="IndexOutOfRangeException">if offset is negative</exception>
        public ShortField(int offset, short value, ref byte[] data)
            : this(offset)
        {
            Set(value, ref data);
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The value.</value>
        public short Value
        {
            get { return _value; }
            set { this._value = value; }
        }
        /// <summary>
        /// set the ShortField's current value and write it to a byte array
        /// </summary>
        /// <param name="value">value to be set</param>
        /// <param name="data">the byte array to write the value to</param>
        /// <exception cref="IndexOutOfRangeException">if the offset is out
        /// of range</exception>
        public void Set(short value, ref byte [] data)
        {
            _value = value;
            WriteToBytes(data);
        }

        /// <summary>
        /// set the value from its offset into an array of bytes
        /// </summary>
        /// <param name="data">the byte array from which the value is to be read</param>
        /// <exception cref="IndexOutOfRangeException">if the offset is out
        /// of range</exception>
        public void ReadFromBytes(byte [] data)
        {
            _value = LittleEndian.GetShort(data, _offset);
        }

        /// <summary>
        /// set the value from an Stream
        /// </summary>
        /// <param name="stream">the Stream from which the value is to be
        /// read</param>
        /// <exception cref="IOException">if an IOException is thrown from reading
        /// the Stream</exception>
        /// <exception cref="BufferUnderrunException">if there is not enough data
        /// available from the Stream</exception>
        public void ReadFromStream(Stream stream)
        {
            _value = LittleEndian.ReadShort(stream);
        }

        /// <summary>
        /// write the value out to an array of bytes at the appropriate
        /// offset
        /// </summary>
        /// <param name="data">the array of bytes to which the value is to be
        /// written</param>
        /// <exception cref="IndexOutOfRangeException">if the offset is out
        /// of range</exception>
        public void WriteToBytes(byte [] data)
        {
            LittleEndian.PutShort(data, _offset, _value);
        }

        /// <summary>
        /// Same as using the constructor <see cref="ShortField"/> with the same
        /// parameter list. Avoid creation of an useless object.
        /// </summary>
        /// <param name="offset">offset of the field within its byte array</param>
        /// <param name="value">the initial value</param>
        /// <param name="data">the byte array to write the value to</param>
        public static void Write(int offset, short value, ref byte[] data)
        {
            LittleEndian.PutShort(data, offset, value);
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            return Convert.ToString(_value, CultureInfo.CurrentCulture);
        }
    }
}
