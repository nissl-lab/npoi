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
using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record
{
    /**
 * DConRef records specify a range in a workbook (internal or external) that serves as a data source
 * for pivot tables or data consolidation.
 *
 * Represents a <code>DConRef</code> Structure
 * <a href="http://msdn.microsoft.com/en-us/library/dd923854(office.12).aspx">[MS-XLS s.
 * 2.4.86]</a>, and the contained <code>DConFile</code> structure
 * <a href="http://msdn.microsoft.com/en-us/library/dd950157(office.12).aspx">
 * [MS-XLS s. 2.5.69]</a>. This in turn contains a <code>XLUnicodeStringNoCch</code>
 * <a href="http://msdn.microsoft.com/en-us/library/dd910585(office.12).aspx">
 * [MS-XLS s. 2.5.296]</a>.
 *
 * <pre>
 *         _______________________________
 *        |          DConRef              |
 *(bytes) +-+-+-+-+-+-+-+-+-+-+...+-+-+-+-+
 *        |    ref    |cch|  stFile   | un|
 *        +-+-+-+-+-+-+-+-+-+-+...+-+-+-+-+
 *                              |
 *                     _________|_____________________
 *                    |DConFile / XLUnicodeStringNoCch|
 *                    +-+-+-+-+-+-+-+-+-+-+-+...+-+-+-+
 *             (bits) |h|   reserved  |      rgb      |
 *                    +-+-+-+-+-+-+-+-+-+-+-+...+-+-+-+
 * </pre>
 * Where
 * <ul>
 * <li><code>DConFile.h = 0x00</code> if the characters in<code>rgb</code> are single byte, and
 * <code>DConFile.h = 0x01</code> if they are double byte. <br/>
 * If they are double byte, then<br/>
 * <ul>
 * <li> If it exists, the length of <code>DConRef.un = 2</code>. Otherwise it is 1.</li>
 * <li> The length of <code>DConFile.rgb = (2 * DConRef.cch)</code>. Otherwise it is equal to
 * <code>DConRef.cch</code></li>.
 * </ul>
 * </li>
 * <li><code>DConRef.rgb</code> starts with <code>0x01</code> if it is an external reference,
 * and with <code>0x02</code> if it is a self-reference.</li>
 * </ul>
 *
 * At the moment this class is read-only.
 *
 * @author Niklas Rehfeld
 */
    public class DConRefRecord : StandardRecord
    {

        /**
         * The id of the record type,
         * <code>sid = {@value}</code>
         */
        public const short sid = 0x0051;
        /**
         * A RefU structure specifying the range of cells if this record is part of an SXTBL.
         * <a href="http://msdn.microsoft.com/en-us/library/dd920420(office.12).aspx">
         * [MS XLS s.2.5.211]</a>
         */
        private int firstRow, lastRow, firstCol, lastCol;
        /**
         * the number of chars in the link
         */
        private int charCount;
        /**
         * the type of characters (single or double byte)
         */
        private int charType;
        /**
         * The link's path string. This is the <code>rgb</code> field of a
         * <code>XLUnicodeStringNoCch</code>. Therefore it will contain at least one leading special
         * character (0x01 or 0x02) and probably other ones.<p/>
         * @see <a href="http://msdn.microsoft.com/en-us/library/dd923491(office.12).aspx">
         * DConFile [MS-XLS s. 2.5.77]</a> and
         * <a href="http://msdn.microsoft.com/en-us/library/dd950157(office.12).aspx">
         * VirtualPath [MS-XLS s. 2.5.69]</a>
         * <p/>
         */
        private byte[] path;
        /**
         * unused bits at the end, must be set to 0.
         */
        private byte[] _unused;

        /**
         * Read constructor.
         *
         * @param data byte array containing a DConRef Record, including the header.
         */
        public DConRefRecord(byte[] data)
        {
            int offset = 0;
            if (!(LittleEndian.GetShort(data, offset) == DConRefRecord.sid))
                throw new RecordFormatException("incompatible sid.");
            offset += LittleEndian.SHORT_SIZE;

            //length = LittleEndian.GetShort(data, offset);
            offset += LittleEndian.SHORT_SIZE;

            firstRow = LittleEndian.GetUShort(data, offset);
            offset += LittleEndian.SHORT_SIZE;
            lastRow = LittleEndian.GetUShort(data, offset);
            offset += LittleEndian.SHORT_SIZE;
            firstCol = LittleEndian.GetUByte(data, offset);
            offset += LittleEndian.BYTE_SIZE;
            lastCol = LittleEndian.GetUByte(data, offset);
            offset += LittleEndian.BYTE_SIZE;
            charCount = LittleEndian.GetUShort(data, offset);
            offset += LittleEndian.SHORT_SIZE;
            if (charCount < 2)
                throw new RecordFormatException("Character count must be >= 2");

            charType = LittleEndian.GetUByte(data, offset);
            offset += LittleEndian.BYTE_SIZE; //7 bits reserved + 1 bit type

            /*
             * bytelength is the length of the string in bytes, which depends on whether the string is
             * made of single- or double-byte chars. This is given by charType, which equals 0 if
             * single-byte, 1 if double-byte.
             */
            int byteLength = charCount * ((charType & 1) + 1);

            path = LittleEndian.GetByteArray(data, offset, byteLength);
            offset += byteLength;

            /*
             * If it's a self reference, the last one or two bytes (depending on char type) are the
             * unused field. Not sure If i need to bother with this...
             */
            if (path[0] == 0x02)
                _unused = LittleEndian.GetByteArray(data, offset, (charType + 1));

        }

        /**
         * Read Constructor.
         *
         * @param inStream RecordInputStream containing a DConRefRecord structure.
         */
        public DConRefRecord(RecordInputStream inStream)
        {
            if (inStream.Sid != sid)
                throw new RecordFormatException("Wrong sid: " + inStream.Sid);

            firstRow = inStream.ReadUShort();
            lastRow = inStream.ReadUShort();
            firstCol = inStream.ReadUByte();
            lastCol = inStream.ReadUByte();

            charCount = inStream.ReadUShort();
            charType = inStream.ReadUByte() & 0x01; //first bit only.

            // byteLength depends on whether we are using single- or double-byte chars.
            int byteLength = charCount * (charType + 1);

            path = new byte[byteLength];
            inStream.ReadFully(path);

            if (path[0] == 0x02)
                _unused = inStream.ReadRemainder();

        }

        /*
         * assuming this wants the number of bytes returned by {@link serialize(LittleEndianOutput)},
         * that is, (length - 4).
         */
        protected override int DataSize
        {
            get
            {
                int sz = 9 + path.Length;
                if (path[0] == 0x02)
                    sz += _unused.Length;
                return sz;
            }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(firstRow);
            out1.WriteShort(lastRow);
            out1.WriteByte(firstCol);
            out1.WriteByte(lastCol);
            out1.WriteShort(charCount);
            out1.WriteByte(charType);
            out1.Write(path);
            if (path[0] == 0x02)
                out1.Write(_unused);
        }

        public override short Sid
        {
            get { return sid; }
        }

        /**
         * @return The first column of the range.
         */
        public int FirstColumn
        {
            get { return firstCol; }
        }

        /**
         * @return The first row of the range.
         */
        public int FirstRow
        {
            get { return firstRow; }
        }

        /**
         * @return The last column of the range.
         */
        public int LastColumn
        {
            get { return lastCol; }
        }

        /**
         * @return The last row of the range.
         */
        public int LastRow
        {
            get { return lastRow; }
        }

        public override String ToString()
        {
            StringBuilder b = new StringBuilder();
            b.Append("[DCONREF]\n");
            b.Append("    .ref\n");
            b.Append("        .firstrow   = ").Append(firstRow).Append("\n");
            b.Append("        .lastrow    = ").Append(lastRow).Append("\n");
            b.Append("        .firstcol   = ").Append(firstCol).Append("\n");
            b.Append("        .lastcol    = ").Append(lastCol).Append("\n");
            b.Append("    .cch            = ").Append(charCount).Append("\n");
            b.Append("    .stFile\n");
            b.Append("        .h          = ").Append(charType).Append("\n");
            b.Append("        .rgb        = ").Append(ReadablePath).Append("\n");
            b.Append("[/DCONREF]\n");

            return b.ToString();
        }

        /**
         *
         * @return raw path byte array.
         */
        public byte[] GetPath()
        {
            return Arrays.CopyOf(path, path.Length);
        }

        /**
         * @return the link's path, with the special characters stripped/replaced. May be null.
         * See MS-XLS 2.5.277 (VirtualPath)
         */
        public String ReadablePath
        {
            get
            {
                if (path != null)
                {
                    //all of the path strings start with either 0x02 or 0x01 followed by zero or
                    //more of 0x01..0x08
                    int offset = 1;
                    while (path[offset] < 0x20 && offset < path.Length)
                    {
                        offset++;
                    }
                    //String out1 = new String(Arrays.CopyOfRange(path, offset, path.Length));
                    String out1 = Encoding.UTF8.GetString(Arrays.CopyOfRange(path, offset, path.Length));
                    //UNC paths have \u0003 chars as path separators.
                    out1 = out1.Replace("\u0003", "/");
                    return out1;
                }
                return null;
            }
        }

        /**
         * Checks if the data source in this reference record is external to this sheet or internal.
         *
         * @return true iff this is an external reference.
         */
        public bool IsExternalRef
        {
            get
            {
                if (path[0] == 0x01)
                    return true;
                return false;
            }
        }
    }
}