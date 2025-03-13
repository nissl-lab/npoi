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

namespace NPOI.HPSF

{


    using NPOI.Util;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    /// <summary>
    /// <para>
    /// </para>
    /// <para>
    /// A property in a <see cref="Section"/> of a <see cref="PropertySet"/>.
    /// </para>
    /// <para>
    /// The property's <c>ID</c> gives the property a meaning
    /// in the context of its <see cref="Section"/>. Each <see cref="Section"/> spans
    /// </para>
    /// <para>
    /// its own name space of property IDs.
    /// </para>
    /// <para>
    /// The property's <c>type</c> determines how its
    /// <c>value</c> is interpreted. For example, if the type is
    /// <see cref="Variant.VT_LPSTR"/> (byte string), the value consists of a
    /// DWord telling how many bytes the string contains. The bytes follow
    /// immediately, including any null bytes that terminate the
    /// string. The type <see cref="Variant.VT_I4"/> denotes a four-byte integer
    /// </para>
    /// <para>
    /// value, <see cref="Variant.VT_FILETIME"/> some date and time (of a file).
    /// </para>
    /// <para>
    /// Please note that not all <see cref="Variant"/> types yet. This might change
    /// over time but largely depends on your feedback so that the POI team knows
    /// which variant types are really needed. So please feel free to submit error
    /// reports or patches for the types you need.
    /// </para>
    /// <para>See <see cref="Section"/></para>
    /// <para>See <see cref="Variant"/></para>
    /// <para>
    /// see <see href="https://msdn.microsoft.com/en-us/library/dd942421.aspx">
    /// [MS-OLEPS]: Object Linking and Embedding (OLE) Property Set Data Structures</see></para>
    /// </summary>
    public class Property
    {

        /// <summary>
        /// The property's ID. */
        /// </summary>
        private long id;

        /// <summary>
        /// The property's type. */
        /// </summary>
        private long type;

        /// <summary>
        /// The property's value. */
        /// </summary>
        private object _value;


        /// <summary>
        /// Creates an empty property. It must be filled using the Set method to be usable.
        /// </summary>
        public Property()
        {
        }

        /// <summary>
        /// Creates a <c>Property</c> as a copy of an existing <c>Property</c>.
        /// </summary>
        /// <param name="p">The property to copy.</param>
        public Property(Property p)
                : this(p.id, p.type, p._value)
        {
            ;
        }

        /// <summary>
        /// Creates a property.
        /// </summary>
        /// <param name="id">the property's ID.</param>
        /// <param name="type">the property's type, see <see cref="Variant"/>.</param>
        /// <param name="value">the property's value. Only certain types are allowed, see
        /// <see cref="Variant"/>.
        /// </param>
        public Property(long id, long type, Object value)
        {
            this.id = id;
            this.type = type;
            this._value = value;
        }

        /// <summary>
        /// Creates a <see cref="Property"/> instance by reading its bytes
        /// from the property Set stream.
        /// </summary>
        /// <param name="id">The property's ID.</param>
        /// <param name="src">The bytes the property Set stream consists of.</param>
        /// <param name="offset">The property's type/value pair's offset in the
        /// section.
        /// </param>
        /// <param name="length">The property's type/value pair's length in bytes.</param>
        /// <param name="codepage">The section's and thus the property's
        /// codepage. It is needed only when reading string values.
        /// </param>
        /// <exception cref="UnsupportedEncodingException">if the specified codepage is not
        /// supported.
        /// </exception>
        public Property(long id, byte[] src, long offset, int length, int codepage)
        {

            this.id = id;

            /*
             * ID 0 is a special case since it specifies a dictionary of
             * property IDs and property names.
             */
            if(id == 0)
            {
                _value = ReadDictionary(src, offset, length, codepage);
                return;
            }

            int o = (int)offset;
            type = LittleEndian.GetUInt(src, o);
            o += LittleEndianConsts.INT_SIZE;

            try
            {
                _value = VariantSupport.Read(src, o, length, (int) type, codepage);
            }
            catch(UnsupportedVariantTypeException ex)
            {
                VariantSupport.WriteUnsupportedTypeMessage(ex);
                _value = ex.Value;
            }
        }



        /// <summary>
        /// get or set the property's ID.
        /// </summary>
        /// <return>ID value</return>
        public long ID
        {
            get { return id; }
            set { id = value; }
        }

        /// <summary>
        /// get or set the property's type.
        /// </summary>
        /// <return>type value</return>
        public long Type
        {
            get { return type; }
            set { type = value; }
        }

        /// <summary>
        /// get or set the property's value.
        /// </summary>
        /// <return>property's value</return>
        public object Value
        {
            get { return _value; }
            set { _value = value; }
        }

        /// <summary>
        /// Reads a dictionary.
        /// </summary>
        /// <param name="src">The byte array containing the bytes making out the dictionary.</param>
        /// <param name="offset">At this offset within <c>src</c> the dictionary
        /// starts.
        /// </param>
        /// <param name="length">The dictionary contains at most this many bytes.</param>
        /// <param name="codepage">The codepage of the string values.</param>
        /// <return>dictonary</return>
        /// <exception cref="UnsupportedEncodingException">if the dictionary's codepage is not
        /// (yet) supported.
        /// </exception>
        protected Dictionary<object, object> ReadDictionary(byte[] src, long offset, int length, int codepage)
        {
            /* Check whether "offset" points into the "src" array". */
            if(offset < 0 || offset > src.Length)
            {
                throw new HPSFRuntimeException
                    ("Illegal offset " + offset + " while HPSF stream contains " +
                     length + " bytes.");
            }
            long o = offset;

            /*
             * Read the number of dictionary entries.
             */
            long nrEntries = LittleEndian.GetUInt(src, (int)o);
            o += LittleEndianConsts.INT_SIZE;

            Dictionary<object, object> m = new Dictionary<object, object>();

            try
            {
                for(int i = 0; i < nrEntries; i++)
                {
                    /* The key. */
                    long id = LittleEndian.GetUInt(src, (int)o);
                    o += LittleEndianConsts.INT_SIZE;

                    /* The value (a string). The length is the either the
                     * number of (two-byte) characters if the character Set is Unicode
                     * or the number of bytes if the character Set is not Unicode.
                     * The length includes terminating 0x00 bytes which we have to strip
                     * off to create a Java string. */
                    long sLength = LittleEndian.GetUInt(src, (int)o);
                    o += LittleEndianConsts.INT_SIZE;

                    /* Read the string. */
                    StringBuilder b = new StringBuilder();
                    switch(codepage)
                    {
                        case -1:
                            /* Without a codepage the length is equal to the number of
                             * bytes. */
                            b.Append(Encoding.GetEncoding("ASCII").GetString(src, (int) o, (int) sLength));
                            break;
                        case CodePageUtil.CP_UNICODE:
                            /* The length is the number of characters, i.e. the number
                             * of bytes is twice the number of the characters. */
                            int nrBytes = (int)(sLength * 2);
                            byte[] h = new byte[nrBytes];
                            for(int i2 = 0; i2 < nrBytes; i2 += 2)
                            {
                                h[i2] = src[o + i2];
                                h[i2 + 1] = src[o + i2 + 1];
                            }
                            b.Append(Encoding.GetEncoding(CodePageUtil.CodepageToEncoding(codepage)).GetString(h, 0, nrBytes));
                            break;
                        default:
                            /* For encodings other than Unicode the length is the number
                             * of bytes. */
                            b.Append(Encoding.GetEncoding(CodePageUtil.CodepageToEncoding(codepage)).GetString(src, (int) o, (int) sLength));
                            break;
                    }

                    /* Strip 0x00 characters from the end of the string: */
                    while(b.Length > 0 && b[b.Length - 1] == 0x00)
                    {
                        b.Length = (b.Length - 1);
                    }
                    if(codepage == CodePageUtil.CP_UNICODE)
                    {
                        if(sLength % 2 == 1)
                        {
                            sLength++;
                        }
                        o += (sLength + sLength);
                    }
                    else
                    {
                        o += sLength;
                    }
                    m.Add(id, b.ToString());
                }
            }
            catch(SystemException)
            {
                //POILogger l = POILogFactory.GetLogger(getClass());
                //l.log(POILogger.WARN,
                //        "The property Set's dictionary contains bogus data. "
                //        + "All dictionary entries starting with the one with ID "
                //        + id + " will be ignored.", ex);
            }
            return m;
        }



        /// <summary>
        /// Returns the property's size in bytes. This is always a multiple of 4.
        /// </summary>
        /// <return>property's size in bytes</return>
        /// 
        /// <exception cref="WritingNotSupportedException">if HPSF does not yet support the
        /// property's variant type.
        /// </exception>
        protected internal int GetSize()

        {
            int length = VariantSupport.GetVariantLength(type);
            if(length >= 0)
            {
                return length; /* Fixed length */
            }
            if(length == -2)
            {
                /* Unknown length */
                throw new WritingNotSupportedException(type, null);
            }

            /* Variable length: */
            int PADDING = 4; /* Pad to multiples of 4. */
            switch((int) type)
            {
                case Variant.VT_LPSTR:
                {
                    int l = ((String)_value).Length + 1;
                    int r = l % PADDING;
                    if(r > 0)
                        l += PADDING - r;
                    length += l;
                    break;
                }
                case Variant.VT_EMPTY:
                    break;
                default:
                    throw new WritingNotSupportedException(type, _value);
            }
            return length;
        }



        /// <summary>
        /// <para>
        /// Compares two properties.
        /// </para>
        /// <para>
        /// Please beware that a property with
        /// ID == 0 is a special case: It does not have a type, and its value is the
        /// section's dictionary. Another special case are strings: Two properties
        /// may have the different types Variant.VT_LPSTR and Variant.VT_LPWSTR;
        /// </para>
        /// </summary>
        /// @see Object#equals(java.lang.Object)
        public override bool Equals(Object o)
        {
            if(o is not Property property)
            {
                return false;
            }

            Object pValue = property.Value;
            long pId = property.ID;
            if(id != pId || (id != 0 && !typesAreEqual(type, property.Type)))
            {
                return false;
            }
            if(_value == null && pValue == null)
            {
                return true;
            }
            if(_value == null || pValue == null)
            {
                return false;
            }

            /* It's clear now that both values are non-null. */
            Type valueClass = _value.GetType();
            Type pValueClass = pValue.GetType();
            if(!(valueClass.IsAssignableFrom(pValueClass)) &&
                !(pValueClass.IsAssignableFrom(valueClass)))
            {
                return false;
            }

            if(_value is byte[] bytes)
            {
                return Arrays.Equals(bytes, (byte[]) pValue);
            }

            return _value.Equals(pValue);
        }



        private bool typesAreEqual(long t1, long t2)
        {
            return (t1 == t2 ||
                (t1 == Variant.VT_LPSTR && t2 == Variant.VT_LPWSTR) ||
                (t2 == Variant.VT_LPSTR && t1 == Variant.VT_LPWSTR));
        }



        /// <summary>
        /// </summary>
        /// @see Object#hashCode()
        public override int GetHashCode()
        {
            long hashCode = 0;
            hashCode += id;
            hashCode += type;
            if(_value != null)
            {
                hashCode += _value.GetHashCode();
            }
            return (int) (hashCode & 0x0ffffffffL);

        }



        /// <summary>
        /// </summary>
        /// @see Object#toString()
        public override String ToString()
        {
            StringBuilder b = new StringBuilder();
            b.Append(GetType().Name);
            b.Append('[');
            b.Append("id: ");
            b.Append(ID);
            b.Append(", type: ");
            b.Append(Type);
            Object value = Value;
            b.Append(", value: ");
            if(value is String s)
            {
                b.Append(s.ToString());
                int l = s.Length;
                byte[] bytes = new byte[l * 2];
                for(int i = 0; i < l; i++)
                {
                    char c = s[i];
                    byte high = (byte)((c & 0x00ff00) >> 8);
                    byte low = (byte)((c & 0x0000ff) >> 0);
                    bytes[i * 2] = high;
                    bytes[i * 2 + 1] = low;
                }
                b.Append(" [");
                if(bytes.Length > 0)
                {
                    String hex = HexDump.Dump(bytes, 0L, 0);
                    b.Append(hex);
                }
                b.Append("]");
            }
            else if(value is byte[] bytes)
            {
                if(bytes.Length > 0)
                {
                    String hex = HexDump.Dump(bytes, 0L, 0);
                    b.Append(hex);
                }
            }
            else
            {
                b.Append(value.ToString());
            }
            b.Append(']');
            return b.ToString();
        }

        /// <summary>
        /// Writes the property to an output stream.
        /// </summary>
        /// <param name="out">The output stream to write to.</param>
        /// <param name="codepage">The codepage to use for writing non-wide strings</param>
        /// <return>number of bytes written to the stream</return>
        /// 
        /// <exception cref="IOException">if an I/O error occurs</exception>
        /// <exception cref="WritingNotSupportedException">if a variant type is to be
        /// written that is not yet supported
        /// </exception>
        public int Write(Stream out1, int codepage)
        {

            int length = 0;
            long variantType = Type;

            /* Ensure that wide strings are written if the codepage is Unicode. */
            if(codepage == CodePageUtil.CP_UNICODE && variantType == Variant.VT_LPSTR)
            {
                variantType = Variant.VT_LPWSTR;
            }

            length += TypeWriter.WriteUIntToStream(out1, (uint) variantType);
            length += VariantSupport.Write(out1, variantType, Value, codepage);
            return length;
        }

    }
}

