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

using System;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.Util;

namespace NPOI.HPSF
{
    /// <summary>
    ///  Represents a class ID (16 bytes). Unlike other little-endian
    ///  type the <see cref="ClassID"/> is not just 16 bytes stored in the wrong
    ///  order. Instead, it is a double word (4 bytes) followed by two
    ///  words (2 bytes each) followed by 8 bytes.
    /// </summary>
    public class ClassID
    {
        public static ClassID OLE10_PACKAGE = new ClassID("{0003000C-0000-0000-C000-000000000046}");
        public static ClassID PPT_SHOW = new ClassID("{64818D10-4F9B-11CF-86EA-00AA00B929E8}");
        public static ClassID XLS_WORKBOOK = new ClassID("{00020841-0000-0000-C000-000000000046}");
        public static ClassID TXT_ONLY = new ClassID("{5e941d80-bf96-11cd-b579-08002b30bfeb}");

        // Excel V3
        public static ClassID EXCEL_V3 = new ClassID("{00030000-0000-0000-C000-000000000046}");
        public static ClassID EXCEL_V3_CHART = new ClassID("{00030001-0000-0000-C000-000000000046}");
        public static ClassID EXCEL_V3_MACRO = new ClassID("{00030002-0000-0000-C000-000000000046}");
        // Excel V5
        public static ClassID EXCEL95 = new ClassID("{00020810-0000-0000-C000-000000000046}");
        public static ClassID EXCEL95_CHART = new ClassID("{00020811-0000-0000-C000-000000000046}");
        // Excel V8
        public static ClassID EXCEL97 = new ClassID("{00020820-0000-0000-C000-000000000046}");
        public static ClassID EXCEL97_CHART = new ClassID("{00020821-0000-0000-C000-000000000046}");
        // Excel V11
        public static ClassID EXCEL2003 = new ClassID("{00020812-0000-0000-C000-000000000046}");
        // Excel V12
        public static ClassID EXCEL2007 = new ClassID("{00020830-0000-0000-C000-000000000046}");
        public static ClassID EXCEL2007_MACRO = new ClassID("{00020832-0000-0000-C000-000000000046}");
        public static ClassID EXCEL2007_XLSB = new ClassID("{00020833-0000-0000-C000-000000000046}");
        // Excel V14
        public static ClassID EXCEL2010 = new ClassID("{00024500-0000-0000-C000-000000000046}");
        public static ClassID EXCEL2010_CHART = new ClassID("{00024505-0014-0000-C000-000000000046}");
        public static ClassID EXCEL2010_ODS = new ClassID("{EABCECDB-CC1C-4A6F-B4E3-7F888A5ADFC8}");

        public static ClassID WORD97 = new ClassID("{00020906-0000-0000-C000-000000000046}");
        public static ClassID WORD95 = new ClassID("{00020900-0000-0000-C000-000000000046}");
        public static ClassID WORD2007 = new ClassID("{F4754C9B-64F5-4B40-8AF4-679732AC0607}");
        public static ClassID WORD2007_MACRO = new ClassID("{18A06B6B-2F3F-4E2B-A611-52BE631B2D22}");

        public static ClassID POWERPOINT97 = new ClassID("{64818D10-4F9B-11CF-86EA-00AA00B929E8}");
        public static ClassID POWERPOINT95 = new ClassID("{EA7BAE70-FB3B-11CD-A903-00AA00510EA3}");
        public static ClassID POWERPOINT2007 = new ClassID("{CF4F55F4-8F87-4D47-80BB-5808164BB3F8}");
        public static ClassID POWERPOINT2007_MACRO = new ClassID("{DC020317-E6E2-4A62-B9FA-B3EFE16626F4}");

        public static ClassID EQUATION30 = new ClassID("{0002CE02-0000-0000-C000-000000000046}");

        /// <summary>
        /// The number of bytes occupied by this object in the byte
        /// stream. */
        /// </summary>
        public const int LENGTH = 16;

        /// <summary>
        /// The bytes making out the class ID in correct order,
        /// i.e. big-endian.
        /// </summary>
        protected byte[] bytes;

        /// <summary>
        /// Creates a <see cref="ClassID"/> and Reads its value from a byte array.
        /// </summary>
        /// <param name="src">The byte array to Read from.</param>
        /// <param name="offset">The offset of the first byte to Read.</param>
        public ClassID(byte[] src, int offset)
        {
            Read(src, offset);
        }


        /// <summary>
        /// Creates a <see cref="ClassID"/> and initializes its value with 0x00 bytes.
        /// </summary>
        public ClassID()
        {
            bytes = new byte[LENGTH];
            for(int i = 0; i < LENGTH; i++)
                bytes[i] = 0x00;
        }

        /**
         * <p>Creates a {@link ClassID} from a human-readable representation of the Class ID in standard 
         * format <code>"{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"</code>.</p>
         * 
         * @param externalForm representation of the Class ID represented by this object.
         */
        public ClassID(String externalForm)
        {
            bytes = new byte[LENGTH];
            String clsStr = Regex.Replace(externalForm, "[{}-]", "");
            for(int i = 0; i < clsStr.Length; i += 2)
            {
                bytes[i / 2] = (byte) Convert.ToInt64(clsStr.Substring(i, 2), 16);
            }
        }

        /// <summary>
        /// Gets the length.
        /// </summary>
        /// <value>The number of bytes occupied by this object in the byte stream.</value>
        public int Length
        {
            get { return LENGTH; }
        }



        /// <summary>
        /// Gets or sets the bytes making out the class ID. They are returned in correct order, i.e. big-endian.
        /// </summary>
        /// <value>the bytes making out the class ID..</value>
        public byte[] Bytes
        {
            get { return bytes; }
            set
            {
                for(int i = 0; i < this.bytes.Length; i++)
                    this.bytes[i] = value[i];
            }
        }

        /// <summary>
        /// Reads the class ID's value from a byte array by turning little-endian into big-endian.
        /// </summary>
        /// <param name="src">The byte array to Read from</param>
        /// <param name="offset">The offset within the </param>
        /// <returns>A byte array containing the class ID.</returns>
        public byte[] Read(byte[] src, int offset)
        {
            bytes = new byte[16];

            /* Read double word. */
            bytes[0] = src[3 + offset];
            bytes[1] = src[2 + offset];
            bytes[2] = src[1 + offset];
            bytes[3] = src[0 + offset];

            /* Read first word. */
            bytes[4] = src[5 + offset];
            bytes[5] = src[4 + offset];

            /* Read second word. */
            bytes[6] = src[7 + offset];
            bytes[7] = src[6 + offset];

            /* Read 8 bytes. */
            for(int i = 8; i < 16; i++)
                bytes[i] = src[i + offset];

            return bytes;
        }



        /// <summary>
        /// Writes the class ID to a byte array in the little-endian format.
        /// </summary>
        /// <param name="dst">The byte array to Write to.</param>
        /// <param name="offset">The offset within the </param>
        /// <exception cref="ArrayTypeMismatchException">if there is not enough room for the class
        /// ID 16 bytes in the byte array After the <var>offset</var> position.
        /// </exception>
        public void Write(byte[] dst, int offset)
        {
            /* Check array size: */
            if(dst.Length < 16)
                throw new ArrayTypeMismatchException
                    ("Destination byte[] must have room for at least 16 bytes, " +
                     "but has a length of only " + dst.Length + ".");
            /* Write double word. */
            dst[0 + offset] = bytes[3];
            dst[1 + offset] = bytes[2];
            dst[2 + offset] = bytes[1];
            dst[3 + offset] = bytes[0];

            /* Write first word. */
            dst[4 + offset] = bytes[5];
            dst[5 + offset] = bytes[4];

            /* Write second word. */
            dst[6 + offset] = bytes[7];
            dst[7 + offset] = bytes[6];

            /* Write 8 bytes. */
            for(int i = 8; i < 16; i++)
                dst[i + offset] = bytes[i];
        }



        /// <summary>
        /// Checks whether this <c>ClassID</c> is equal to another
        /// object.
        /// </summary>
        /// <param name="o">the object to compare this <c>PropertySet</c> with</param>
        /// <return><c>true</c> if the objects are equal, else
        /// <c>false</c>.
        /// </return>
        public override bool Equals(Object o)
        {
            return o is ClassID cid && Arrays.Equals(bytes, cid.bytes);
        }

        /**
         * Checks whether this {@code ClassID} is equal to another ClassID with inverted endianess,
         * because there are apparently not only version 1 GUIDs (aka "network" with big-endian encoding),
         * but also version 2 GUIDs (aka "native" with little-endian encoding) out there.
         *
         * @param o the object to compare this {@code ClassID} with
         * @return {@code true} if the objects are equal, else {@code false}.
         */
        public bool EqualsInverted(ClassID o)
        {
            return
                o.bytes[0] == bytes[3] &&
                o.bytes[1] == bytes[2] &&
                o.bytes[2] == bytes[1] &&
                o.bytes[3] == bytes[0] &&
                o.bytes[4] == bytes[5] &&
                o.bytes[5] == bytes[4] &&
                o.bytes[6] == bytes[7] &&
                o.bytes[7] == bytes[6] &&
                o.bytes[8] == bytes[8] &&
                o.bytes[9] == bytes[9] &&
                o.bytes[10] == bytes[10] &&
                o.bytes[11] == bytes[11] &&
                o.bytes[12] == bytes[12] &&
                o.bytes[13] == bytes[13] &&
                o.bytes[14] == bytes[14] &&
                o.bytes[15] == bytes[15]
                ;
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        /// <summary>
        /// Returns a human-Readable representation of the Class ID in standard
        /// format <c>"{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"</c>.
        /// </summary>
        /// <returns>
        /// A String representation of the Class ID represented by this object..
        /// </returns>
        public override String ToString()
        {
            StringBuilder sbClassId = new StringBuilder(38);
            sbClassId.Append('{');
            for(int i = 0; i < 16; i++)
            {
                sbClassId.Append(HexDump.ToHex(bytes[i]));
                if(i == 3 || i == 5 || i == 7 || i == 9)
                    sbClassId.Append('-');
            }
            sbClassId.Append('}');
            return sbClassId.ToString();
        }

    }
}
