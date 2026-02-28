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
     * Footnote Reference Descriptor (FRD).
     * <p>
     * Class and fields descriptions are quoted from Microsoft Office Word 97-2007
     * Binary File Format (.doc) Specification
     * 
     * NOTE: This source is automatically generated please do not modify this file.
     * Either subclass or remove the record in src/types/defInitions.
     * 
     * @author Sergey Vladimirov; according to Microsoft Office Word 97-2007 Binary
     *         File Format (.doc) Specification
     */

    public abstract class FRDAbstractType:BaseObject
    {

        protected short field_1_nAuto;

        protected FRDAbstractType()
        {
        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_nAuto = LittleEndian.GetShort(data, 0x0 + offset);
        }

        public void Serialize(byte[] data, int offset)
        {
            LittleEndian.PutShort(data, 0x0 + offset, field_1_nAuto);
        }

        /**
         * Size of record
         */
        public static int GetSize()
        {
            return 0 + 2;
        }

        public override String ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("[FRD]\n");
            builder.Append("    .nAuto                = ");
            builder.Append(" (").Append(GetNAuto()).Append(" )\n");

            builder.Append("[/FRD]\n");
            return builder.ToString();
        }

        /**
         * If > 0, the note is an automatically numbered note, otherwise it has a
         * custom mark.
         */
        public short GetNAuto()
        {
            return field_1_nAuto;
        }

        /**
         * If > 0, the note is an automatically numbered note, otherwise it has a
         * custom mark.
         */
        public void SetNAuto(short field_1_nAuto)
        {
            this.field_1_nAuto = field_1_nAuto;
        }

    }
}


