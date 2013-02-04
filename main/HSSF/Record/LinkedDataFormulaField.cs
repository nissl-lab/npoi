
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using System.Text;

    using NPOI.SS.Formula.PTG;

    /**
     * Not implemented yet. May commit it anyway just so people can see
     * where I'm heading.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class LinkedDataFormulaField
    {
        Ptg[] formulaTokens;

        public int Size
        {
            get
            {
                return 2 + Ptg.GetEncodedSize(formulaTokens);
            }
        }

        public int FillField(RecordInputStream in1)
        {
            short tokenSize = in1.ReadShort();
            formulaTokens = Ptg.ReadTokens(tokenSize, in1);

            return tokenSize + 2;
        }

        public void toString(StringBuilder buffer)
        {
            for (int k = 0; k < formulaTokens.Length; k++)
            {
                buffer.Append("Formula ")
                        .Append(k)
                        .Append("=")
                        .Append(formulaTokens[k].ToString())
                        .Append("\n");
                        //.Append(((Ptg)formulaTokens[k]).ToDebugString())
                        //.Append("\n");
            }
        }

        public override String ToString()
        {
            StringBuilder b = new StringBuilder();
            toString(b);
            return b.ToString();
        }

        public int SerializeField(int offset, byte[] data)
        {
            int size = Size;
            LittleEndian.PutShort(data, offset, (short)(size - 2));
            int pos = offset + 2;
            pos += Ptg.SerializePtgs(formulaTokens, data, pos);
            return size;
        }


        public Ptg[] FormulaTokens
        {
            get
            {
                return (Ptg[])this.formulaTokens.Clone();
            }
            set 
            {
                this.formulaTokens = (Ptg[])value.Clone();
            }
        }


        public LinkedDataFormulaField Copy()
        {
            LinkedDataFormulaField result = new LinkedDataFormulaField();

            result.formulaTokens = this.FormulaTokens;
            return result;
        }
    }
}