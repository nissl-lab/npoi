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

namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.Util;
    using System.Text;


    /**
     * This structure appears as part of an Obj record that represents image display properties.
     */
    public class FtPioGrbitSubRecord : SubRecord
    {
        public const short sid = 0x08;
        public const short length = 0x02;

        /**
         * A bit that specifies whether the picture's aspect ratio is preserved when rendered in 
         * different views (Normal view, Page Break Preview view, Page Layout view and printing).
         */
        public static int AUTO_PICT_BIT = 1 << 0;

        /**
         * A bit that specifies whether the pictFmla field of the Obj record that Contains 
         * this FtPioGrbit specifies a DDE reference.
         */
        public static int DDE_BIT = 1 << 1;

        /**
         * A bit that specifies whether this object is expected to be updated on print to
         * reflect the values in the cell associated with the object.
         */
        public static int PRINT_CALC_BIT = 1 << 2;

        /**
         * A bit that specifies whether the picture is displayed as an icon.
         */
        public static int ICON_BIT = 1 << 3;

        /**
         * A bit that specifies whether this object is an ActiveX control.
         * It MUST NOT be the case that both fCtl and fDde are equal to 1.
         */
        public static int CTL_BIT = 1 << 4;

        /**
         * A bit that specifies whether the object data are stored in an
         * embedding storage (= 0) or in the controls stream (ctls) (= 1).
         */
        public static int PRSTM_BIT = 1 << 5;

        /**
         * A bit that specifies whether this is a camera picture.
         */
        public static int CAMERA_BIT = 1 << 7;

        /**
         * A bit that specifies whether this picture's size has been explicitly Set.
         * 0 = picture size has been explicitly Set, 1 = has not been Set
         */
        public static int DEFAULT_SIZE_BIT = 1 << 8;

        /**
         * A bit that specifies whether the OLE server for the object is called
         * to load the object's data automatically when the parent workbook is opened.
         */
        public static int AUTO_LOAD_BIT = 1 << 9;


        private short flags = 0;

        /**
         * Construct a new <code>FtPioGrbitSubRecord</code> and
         * fill its data with the default values
         */
        public FtPioGrbitSubRecord()
        {
        }

        public FtPioGrbitSubRecord(ILittleEndianInput in1, int size)
        {
            if (size != length)
            {
                throw new RecordFormatException("Unexpected size (" + size + ")");
            }
            flags = in1.ReadShort();
        }

        /**
         * Use one of the bitmasks MANUAL_ADVANCE_BIT ... CURSOR_VISIBLE_BIT
         * @param bitmask
         * @param enabled
         */
        public void SetFlagByBit(int bitmask, bool enabled)
        {
            if (enabled)
            {
                flags |= (short)bitmask;
            }
            else
            {
                flags &= (short)(0xFFFF ^ bitmask);
            }
        }

        public bool GetFlagByBit(int bitmask)
        {
            return ((flags & bitmask) != 0);
        }

        /**
         * Convert this record to string.
         * Used by BiffViewer and other utilities.
         */
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[FtPioGrbit ]\n");
            buffer.Append("  size     = ").Append(length).Append("\n");
            buffer.Append("  flags    = ").Append(HexDump.ToHex(flags)).Append("\n");
            buffer.Append("[/FtPioGrbit ]\n");
            return buffer.ToString();
        }

        /**
         * Serialize the record data into the supplied array of bytes
         *
         * @param out the stream to serialize into
         */
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(length);
            out1.WriteShort(flags);
        }

        public override int DataSize
        {
            get
            {
                return length;
            }
        }

        /**
         * @return id of this record.
         */
        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public override Object Clone()
        {
            FtPioGrbitSubRecord rec = new FtPioGrbitSubRecord();
            rec.flags = this.flags;
            return rec;
        }

        public short Flags
        {
            get
            {
                return flags;
            }
            set
            {
                this.flags = value;
            }
        }
    }
}