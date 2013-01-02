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
     * Field Descriptor (FLD).
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
    public abstract class FLDAbstractType : BaseObject
    {

        protected byte field_1_chHolder;
        private static BitField ch = new BitField(0x1f);
        private static BitField reserved = new BitField(0xe0);
        protected byte field_2_flt;
        private static BitField fDiffer = new BitField(0x01);
        private static BitField fZombieEmbed = new BitField(0x02);
        private static BitField fResultDirty = new BitField(0x04);
        private static BitField fResultEdited = new BitField(0x08);
        private static BitField fLocked = new BitField(0x10);
        private static BitField fPrivateResult = new BitField(0x20);
        private static BitField fNested = new BitField(0x40);
        private static BitField fHasSep = new BitField(0x40);

        public FLDAbstractType()
        {

        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_chHolder = data[0x0 + offset];
            field_2_flt = data[0x1 + offset];
        }

        public void Serialize(byte[] data, int offset)
        {
            data[0x0 + offset] = field_1_chHolder;
            data[0x1 + offset] = field_2_flt;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FLD]\n");

            buffer.Append("    .chHolder             = ");
            buffer.Append(" (").Append(GetChHolder()).Append(" )\n");
            buffer.Append("         .ch                       = ")
                    .Append(GetCh()).Append('\n');
            buffer.Append("         .reserved                 = ")
                    .Append(GetReserved()).Append('\n');

            buffer.Append("    .flt                  = ");
            buffer.Append(" (").Append(GetFlt()).Append(" )\n");
            buffer.Append("         .fDiffer                  = ")
                    .Append(IsFDiffer()).Append('\n');
            buffer.Append("         .fZombieEmbed             = ")
                    .Append(IsFZombieEmbed()).Append('\n');
            buffer.Append("         .fResultDirty             = ")
                    .Append(IsFResultDirty()).Append('\n');
            buffer.Append("         .fResultEdited            = ")
                    .Append(IsFResultEdited()).Append('\n');
            buffer.Append("         .fLocked                  = ")
                    .Append(IsFLocked()).Append('\n');
            buffer.Append("         .fPrivateResult           = ")
                    .Append(IsFPrivateResult()).Append('\n');
            buffer.Append("         .fNested                  = ")
                    .Append(IsFNested()).Append('\n');
            buffer.Append("         .fHasSep                  = ")
                    .Append(IsFHasSep()).Append('\n');

            buffer.Append("[/FLD]\n");
            return buffer.ToString();
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public static int GetSize()
        {
            return 4 + +1 + 1;
        }

        /**
         * ch field holder (along with reserved bits).
         */
        public byte GetChHolder()
        {
            return field_1_chHolder;
        }

        /**
         * ch field holder (along with reserved bits).
         */
        public void SetChHolder(byte field_1_chHolder)
        {
            this.field_1_chHolder = field_1_chHolder;
        }

        /**
         * Field type when ch == 19 OR field flags when ch == 21 .
         */
        public byte GetFlt()
        {
            return field_2_flt;
        }

        /**
         * Field type when ch == 19 OR field flags when ch == 21 .
         */
        public void SetFlt(byte field_2_flt)
        {
            this.field_2_flt = field_2_flt;
        }

        /**
         * Sets the ch field value. Type of field boundary the FLD describes: 19 --
         * field begin mark, 20 -- field separation mark; 21 -- field end mark
         */
        public void SetCh(byte value)
        {
            field_1_chHolder = (byte)ch.SetValue(field_1_chHolder, value);

        }

        /**
         * Type of field boundary the FLD describes: 19 -- field begin mark, 20 --
         * field separation mark; 21 -- field end mark
         * 
         * @return the ch field value.
         */
        public byte GetCh()
        {
            return (byte)ch.GetValue(field_1_chHolder);

        }

        /**
         * Sets the reserved field value. Reserved
         */
        public void SetReserved(byte value)
        {
            field_1_chHolder = (byte)reserved.SetValue(field_1_chHolder, value);

        }

        /**
         * Reserved
         * 
         * @return the reserved field value.
         */
        public byte GetReserved()
        {
            return (byte)reserved.GetValue(field_1_chHolder);

        }

        /**
         * Sets the fDiffer field value. Ignored for saved file
         */
        public void SetFDiffer(bool value)
        {
            field_2_flt = (byte)fDiffer.SetBoolean(field_2_flt, value);

        }

        /**
         * Ignored for saved file
         * 
         * @return the fDiffer field value.
         */
        public bool IsFDiffer()
        {
            return fDiffer.IsSet(field_2_flt);

        }

        /**
         * Sets the fZombieEmbed field value. ==1 when result still believes this
         * field is an EMBED or LINK field
         */
        public void SetFZombieEmbed(bool value)
        {
            field_2_flt = (byte)fZombieEmbed.SetBoolean(field_2_flt, value);

        }

        /**
         * ==1 when result still believes this field is an EMBED or LINK field
         * 
         * @return the fZombieEmbed field value.
         */
        public bool IsFZombieEmbed()
        {
            return fZombieEmbed.IsSet(field_2_flt);

        }

        /**
         * Sets the fResultDirty field value. ==1 when user has edited or formatted
         * the result. == 0 otherwise
         */
        public void SetFResultDirty(bool value)
        {
            field_2_flt = (byte)fResultDirty.SetBoolean(field_2_flt, value);

        }

        /**
         * ==1 when user has edited or formatted the result. == 0 otherwise
         * 
         * @return the fResultDirty field value.
         */
        public bool IsFResultDirty()
        {
            return fResultDirty.IsSet(field_2_flt);

        }

        /**
         * Sets the fResultEdited field value. ==1 when user has inserted text into
         * or deleted text from the result
         */
        public void SetFResultEdited(bool value)
        {
            field_2_flt = (byte)fResultEdited.SetBoolean(field_2_flt, value);

        }

        /**
         * ==1 when user has inserted text into or deleted text from the result
         * 
         * @return the fResultEdited field value.
         */
        public bool IsFResultEdited()
        {
            return fResultEdited.IsSet(field_2_flt);

        }

        /**
         * Sets the fLocked field value. ==1 when field is locked from recalculation
         */
        public void SetFLocked(bool value)
        {
            field_2_flt = (byte)fLocked.SetBoolean(field_2_flt, value);

        }

        /**
         * ==1 when field is locked from recalculation
         * 
         * @return the fLocked field value.
         */
        public bool IsFLocked()
        {
            return fLocked.IsSet(field_2_flt);

        }

        /**
         * Sets the fPrivateResult field value. ==1 whenever the result of the field
         * is never to be Shown
         */
        public void SetFPrivateResult(bool value)
        {
            field_2_flt = (byte)fPrivateResult.SetBoolean(field_2_flt, value);

        }

        /**
         * ==1 whenever the result of the field is never to be Shown
         * 
         * @return the fPrivateResult field value.
         */
        public bool IsFPrivateResult()
        {
            return fPrivateResult.IsSet(field_2_flt);

        }

        /**
         * Sets the fNested field value. ==1 when field is nested within another
         * field
         */
        public void SetFNested(bool value)
        {
            field_2_flt = (byte)fNested.SetBoolean(field_2_flt, value);

        }

        /**
         * ==1 when field is nested within another field
         * 
         * @return the fNested field value.
         */
        public bool IsFNested()
        {
            return fNested.IsSet(field_2_flt);

        }

        /**
         * Sets the fHasSep field value. ==1 when field has a field separator
         */
        public void SetFHasSep(bool value)
        {
            field_2_flt = (byte)fHasSep.SetBoolean(field_2_flt, value);

        }

        /**
         * ==1 when field has a field separator
         * 
         * @return the fHasSep field value.
         */
        public bool IsFHasSep()
        {
            return fHasSep.IsSet(field_2_flt);

        }
    }

}

