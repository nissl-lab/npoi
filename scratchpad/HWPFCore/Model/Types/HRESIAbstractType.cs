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

using System.Text;
using System;
namespace NPOI.HWPF.Model.Types
{

    /**
     * Hyphenation (HRESI).
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

    public abstract class HRESIAbstractType:BaseObject
    {

        protected byte field_1_hres;
        public static byte HRES_NO = 0;
        public static byte HRES_NORMAL = 1;
        public static byte HRES_ADD_LETTER_BEFORE = 2;
        public static byte HRES_CHANGE_LETTER_BEFORE = 3;
        public static byte HRES_DELETE_LETTER_BEFORE = 4;
        public static byte HRES_CHANGE_LETTER_AFTER = 5;
        public static byte HRES_DELETE_BEFORE_CHANGE_BEFORE = 6;
        protected byte field_2_chHres;

        protected HRESIAbstractType()
        {
        }

        protected void FillFields(byte[] data, int offset)
        {
            field_1_hres = data[0x0 + offset];
            field_2_chHres = data[0x1 + offset];
        }

        public void Serialize(byte[] data, int offset)
        {
            data[0x0 + offset] = field_1_hres;
            data[0x1 + offset] = field_2_chHres;
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        public static int GetSize()
        {
            return 4 + +1 + 1;
        }

        public override String ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("[HRESI]\n");
            builder.Append("    .hres                 = ");
            builder.Append(" (").Append(GetHres()).Append(" )\n");
            builder.Append("    .chHres               = ");
            builder.Append(" (").Append(GetChHres()).Append(" )\n");

            builder.Append("[/HRESI]\n");
            return builder.ToString();
        }

        /**
         * Hyphenation rule.
         *
         * @return One of 
         * <li>{@link #HRES_NO}
         * <li>{@link #HRES_NORMAL}
         * <li>{@link #HRES_ADD_LETTER_BEFORE}
         * <li>{@link #HRES_CHANGE_LETTER_BEFORE}
         * <li>{@link #HRES_DELETE_LETTER_BEFORE}
         * <li>{@link #HRES_CHANGE_LETTER_AFTER}
         * <li>{@link #HRES_DELETE_BEFORE_CHANGE_BEFORE}
         */
        public byte GetHres()
        {
            return field_1_hres;
        }

        /**
         * Hyphenation rule.
         *
         * @param field_1_hres
         *        One of 
         * <li>{@link #HRES_NO}
         * <li>{@link #HRES_NORMAL}
         * <li>{@link #HRES_ADD_LETTER_BEFORE}
         * <li>{@link #HRES_CHANGE_LETTER_BEFORE}
         * <li>{@link #HRES_DELETE_LETTER_BEFORE}
         * <li>{@link #HRES_CHANGE_LETTER_AFTER}
         * <li>{@link #HRES_DELETE_BEFORE_CHANGE_BEFORE}
         */
        public void SetHres(byte field_1_hres)
        {
            this.field_1_hres = field_1_hres;
        }

        /**
         * The character that will be used to add or change a letter when hres is 2, 3, 5 or 6.
         */
        public byte GetChHres()
        {
            return field_2_chHres;
        }

        /**
         * The character that will be used to add or change a letter when hres is 2, 3, 5 or 6.
         */
        public void SetChHres(byte field_2_chHres)
        {
            this.field_2_chHres = field_2_chHres;
        }

    }
}
