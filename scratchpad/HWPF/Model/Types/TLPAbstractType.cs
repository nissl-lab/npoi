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

using NPOI.Util;
using System.Text;
using System;
namespace NPOI.HWPF.Model.Types
{

    /**
     * Table Autoformat Look sPecifier (TLP).
     * <p>
     * Class and fields descriptions are quoted from Microsoft Office Word 97-2007
     * Binary File Format
     * 
     * NOTE: This source is automatically generated please do not modify this file.
     * Either subclass or remove the record in src/records/defInitions.
     * 
     * @author Sergey Vladimirov; according to Microsoft Office Word 97-2007 Binary
     *         File Format Specification [*.doc]
     */

    public abstract class TLPAbstractType : BaseObject
    {

        protected short field_1_itl;
        protected byte field_2_tlp_flags;
        private static BitField fBorders = new BitField(0x0001);
        private static BitField fShading = new BitField(0x0002);
        private static BitField fFont = new BitField(0x0004);
        private static BitField fColor = new BitField(0x0008);
        private static BitField fBestFit = new BitField(0x0010);
        private static BitField fHdrRows = new BitField(0x0020);
        private static BitField fLastRow = new BitField(0x0040);

        public TLPAbstractType()
        {

        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_itl = LittleEndian.GetShort(data, 0x0 + offset);
            field_2_tlp_flags = data[0x2 + offset];
        }

        public void Serialize(byte[] data, int offset)
        {
            LittleEndian.PutShort(data, 0x0 + offset, field_1_itl);
            data[0x2 + offset] = field_2_tlp_flags;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[TLP]\n");

            buffer.Append("    .itl                  = ");
            buffer.Append(" (").Append(GetItl()).Append(" )\n");

            buffer.Append("    .tlp_flags            = ");
            buffer.Append(" (").Append(GetTlp_flags()).Append(" )\n");
            buffer.Append("         .fBorders                 = ")
                    .Append(IsFBorders()).Append('\n');
            buffer.Append("         .fShading                 = ")
                    .Append(IsFShading()).Append('\n');
            buffer.Append("         .fFont                    = ")
                    .Append(IsFFont()).Append('\n');
            buffer.Append("         .fColor                   = ")
                    .Append(IsFColor()).Append('\n');
            buffer.Append("         .fBestFit                 = ")
                    .Append(IsFBestFit()).Append('\n');
            buffer.Append("         .fHdrRows                 = ")
                    .Append(IsFHdrRows()).Append('\n');
            buffer.Append("         .fLastRow                 = ")
                    .Append(IsFLastRow()).Append('\n');

            buffer.Append("[/TLP]\n");
            return buffer.ToString();
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public int GetSize()
        {
            return 4 + +2 + 1;
        }

        /**
         * Get the itl field for the TLP record.
         */
        public short GetItl()
        {
            return field_1_itl;
        }

        /**
         * Set the itl field for the TLP record.
         */
        public void SetItl(short field_1_itl)
        {
            this.field_1_itl = field_1_itl;
        }

        /**
         * Get the tlp_flags field for the TLP record.
         */
        public byte GetTlp_flags()
        {
            return field_2_tlp_flags;
        }

        /**
         * Set the tlp_flags field for the TLP record.
         */
        public void SetTlp_flags(byte field_2_tlp_flags)
        {
            this.field_2_tlp_flags = field_2_tlp_flags;
        }

        /**
         * Sets the fBorders field value. When == 1, use the border properties from
         * the selected table look
         */
        public void SetFBorders(bool value)
        {
            field_2_tlp_flags = (byte)fBorders.SetBoolean(field_2_tlp_flags,
                    value);

        }

        /**
         * When == 1, use the border properties from the selected table look
         * 
         * @return the fBorders field value.
         */
        public bool IsFBorders()
        {
            return fBorders.IsSet(field_2_tlp_flags);

        }

        /**
         * Sets the fShading field value. When == 1, use the shading properties from
         * the selected table look
         */
        public void SetFShading(bool value)
        {
            field_2_tlp_flags = (byte)fShading.SetBoolean(field_2_tlp_flags,
                    value);

        }

        /**
         * When == 1, use the shading properties from the selected table look
         * 
         * @return the fShading field value.
         */
        public bool IsFShading()
        {
            return fShading.IsSet(field_2_tlp_flags);

        }

        /**
         * Sets the fFont field value. When == 1, use the font from the selected
         * table look
         */
        public void SetFFont(bool value)
        {
            field_2_tlp_flags = (byte)fFont.SetBoolean(field_2_tlp_flags, value);

        }

        /**
         * When == 1, use the font from the selected table look
         * 
         * @return the fFont field value.
         */
        public bool IsFFont()
        {
            return fFont.IsSet(field_2_tlp_flags);

        }

        /**
         * Sets the fColor field value. When == 1, use the color from the selected
         * table look
         */
        public void SetFColor(bool value)
        {
            field_2_tlp_flags = (byte)fColor.SetBoolean(field_2_tlp_flags, value);

        }

        /**
         * When == 1, use the color from the selected table look
         * 
         * @return the fColor field value.
         */
        public bool IsFColor()
        {
            return fColor.IsSet(field_2_tlp_flags);

        }

        /**
         * Sets the fBestFit field value. When == 1, do best fit from the selected
         * table look
         */
        public void SetFBestFit(bool value)
        {
            field_2_tlp_flags = (byte)fBestFit.SetBoolean(field_2_tlp_flags,
                    value);

        }

        /**
         * When == 1, do best fit from the selected table look
         * 
         * @return the fBestFit field value.
         */
        public bool IsFBestFit()
        {
            return fBestFit.IsSet(field_2_tlp_flags);

        }

        /**
         * Sets the fHdrRows field value. When == 1, apply properties from the
         * selected table look to the header rows in the table
         */
        public void SetFHdrRows(bool value)
        {
            field_2_tlp_flags = (byte)fHdrRows.SetBoolean(field_2_tlp_flags,
                    value);

        }

        /**
         * When == 1, apply properties from the selected table look to the header
         * rows in the table
         * 
         * @return the fHdrRows field value.
         */
        public bool IsFHdrRows()
        {
            return fHdrRows.IsSet(field_2_tlp_flags);

        }

        /**
         * Sets the fLastRow field value. When == 1, apply properties from the
         * selected table look to the last row in the table
         */
        public void SetFLastRow(bool value)
        {
            field_2_tlp_flags = (byte)fLastRow.SetBoolean(field_2_tlp_flags,
                    value);

        }

        /**
         * When == 1, apply properties from the selected table look to the last row
         * in the table
         * 
         * @return the fLastRow field value.
         */
        public bool IsFLastRow()
        {
            return fLastRow.IsSet(field_2_tlp_flags);

        }

    }
}

