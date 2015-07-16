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

namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using NPOI.Util;
    using NPOI.SS.Formula;

    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Constant;


    /**
     * EXTERNALNAME<p/>
     * 
     * @author Josh Micich
     */
    public class ExternalNameRecord : StandardRecord
    {

        public const short sid = 0x23; // as per BIFF8. (some old versions used 0x223)

        private const int OPT_BUILTIN_NAME = 0x0001;
        private const int OPT_AUTOMATIC_LINK = 0x0002; // m$ doc calls this fWantAdvise 
        private const int OPT_PICTURE_LINK = 0x0004;
        private const int OPT_STD_DOCUMENT_NAME = 0x0008;
        private const int OPT_OLE_LINK = 0x0010;
        //	private const int OPT_CLIP_FORMAT_MASK      = 0x7FE0;
        private const int OPT_ICONIFIED_PICTURE_LINK = 0x8000;


        private short field_1_option_flag;
        private short field_2_ixals;
        private short field_3_not_used;
        private String field_4_name;
        private Formula field_5_name_definition; 

        /**
         * 'rgoper' / 'Last received results of the DDE link'
         * (seems to be only applicable to DDE links)<br/>
         * Logically this is a 2-D array, which has been flattened into 1-D array here.
         */
        private Object[] _ddeValues;
        /**
         * (logical) number of columns in the {@link #_ddeValues} array
         */
        private int _nColumns;
        /**
         * (logical) number of rows in the {@link #_ddeValues} array
         */
        private int _nRows;
        public ExternalNameRecord()
        {
            field_2_ixals = 0;
        }
        public ExternalNameRecord(RecordInputStream in1)
        {
            field_1_option_flag = in1.ReadShort();
            field_2_ixals = in1.ReadShort();
            field_3_not_used = in1.ReadShort();
            int numChars = in1.ReadUByte();
            field_4_name = StringUtil.ReadUnicodeString(in1, numChars);

            // the record body can take different forms.
            // The form is dictated by the values of 3-th and 4-th bits in field_1_option_flag
            if (!IsOLELink && !IsStdDocumentNameIdentifier)
            {
                // another switch: the fWantAdvise bit specifies whether the body describes
                // an external defined name or a DDE data item
                if (IsAutomaticLink)
                {
                    if (in1.Available() > 0)
                    {
                        //body specifies DDE data item
                        int nColumns = in1.ReadUByte() + 1;
                        int nRows = in1.ReadShort() + 1;

                        int totalCount = nRows * nColumns;
                        _ddeValues = ConstantValueParser.Parse(in1, totalCount);
                        _nColumns = nColumns;
                        _nRows = nRows;
                    }
                }
                else
                {
                    //body specifies an external defined name
                    int formulaLen = in1.ReadUShort();
                    field_5_name_definition = Formula.Read(formulaLen, in1);
                }
            }
        }

        /**
         * Convenience Function to determine if the name Is a built-in name
         */
        public bool IsBuiltInName
        {
            get
            {
                return (field_1_option_flag & OPT_BUILTIN_NAME) != 0;
            }
        }
        /**
         * For OLE and DDE, links can be either 'automatic' or 'manual'
         */
        public bool IsAutomaticLink
        {
            get { return (field_1_option_flag & OPT_AUTOMATIC_LINK) != 0; }
        }
        /**
         * only for OLE and DDE
         */
        public bool IsPicureLink
        {
            get { return (field_1_option_flag & OPT_PICTURE_LINK) != 0; }
        }
        /**
         * DDE links only. If <c>true</c>, this denotes the 'StdDocumentName'
         */
        public bool IsStdDocumentNameIdentifier
        {
            get { return (field_1_option_flag & OPT_STD_DOCUMENT_NAME) != 0; }
        }
        public bool IsOLELink
        {
            get { return (field_1_option_flag & OPT_OLE_LINK) != 0; }
        }
        public bool IsIconifiedPictureLink
        {
            get { return (field_1_option_flag & OPT_ICONIFIED_PICTURE_LINK) != 0; }
        }
        public short Ix
        {
            get
            {
                return field_2_ixals;
            }
            set
            {
                field_2_ixals = value;
            }
        }

        /**
         * @return the standard String representation of this name
         */
        public String Text
        {
            get { return field_4_name; }
            set { field_4_name = value; }
        }

        public Ptg[] GetParsedExpression()
        {
            return Formula.GetTokens(field_5_name_definition);
        }
        public void SetParsedExpression(Ptg[] ptgs)
        {
            field_5_name_definition = Formula.Create(ptgs);
        }
        protected override int DataSize
        {
            get
            {
                int result = 2 + 4;  // short and int
                result += StringUtil.GetEncodedSize(field_4_name) - 1; //size is byte, not short 

                if (!IsOLELink && !IsStdDocumentNameIdentifier)
                {
                    if (IsAutomaticLink)
                    {
                        if (_ddeValues != null)
                        {
                            result += 3; // byte, short
                            result += ConstantValueParser.GetEncodedSize(_ddeValues);
                        }
                    }
                    else
                    {
                        result += field_5_name_definition.EncodedSize;
                    }
                }
                return result;
            }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            //out1.WriteShort(field_1_option_flag);
            //out1.WriteShort(field_2_ixals);
            //out1.WriteShort(field_3_not_used);
            //int nameLen = field_4_name.Length;
            //out1.WriteShort(nameLen);
            //StringUtil.PutCompressedUnicode(field_4_name, out1);
            //if (HasFormula)
            //{
            //    field_5_name_definition.Serialize(out1);
            //}
            //else
            //{
            //    if (_ddeValues != null)
            //    {
            //        out1.WriteByte(_nColumns - 1);
            //        out1.WriteShort(_nRows - 1);
            //        ConstantValueParser.Encode(out1, _ddeValues);
            //    }
            //}
            out1.WriteShort(field_1_option_flag);
            out1.WriteShort(field_2_ixals);
            out1.WriteShort(field_3_not_used);

            out1.WriteByte(field_4_name.Length);
            StringUtil.WriteUnicodeStringFlagAndData(out1, field_4_name);

            if (!IsOLELink && !IsStdDocumentNameIdentifier)
            {
                if (IsAutomaticLink)
                {
                    if (_ddeValues != null)
                    {
                        out1.WriteByte(_nColumns - 1);
                        out1.WriteShort(_nRows - 1);
                        ConstantValueParser.Encode(out1, _ddeValues);
                    }
                }
                else
                {
                    field_5_name_definition.Serialize(out1);
                }
            }
        }


        //public override int RecordSize
        //{
        //    get { return 4 + DataSize; }
        //}

        /*
         * Makes better error messages (while HasFormula() Is not reliable) 
         * Remove this when HasFormula() Is stable.
         */
        private Exception ReadFail(String msg)
        {
            String fullMsg = msg + " fields: (option=" + field_1_option_flag + " index=" + field_2_ixals
            + " not_used=" + field_3_not_used + " name='" + field_4_name + "')";
            return new Exception(fullMsg);
        }

        private bool HasFormula
        {
            get
            {
                // TODO - determine exact conditions when formula Is present
                //if (false)
                //{
                //    // "Microsoft Office Excel 97-2007 Binary File Format (.xls) Specification"
                //    // m$'s document suggests logic like this, but bugzilla 44774 att 21790 seems to disagree
                //    if (IsStdDocumentNameIdentifier)
                //    {
                //        if (IsOLELink)
                //        {
                //            // seems to be not possible according to m$ document
                //            throw new InvalidOperationException(
                //                    "flags (std-doc-name and ole-link) cannot be true at the same time");
                //        }
                //        return false;
                //    }
                //    if (IsOLELink)
                //    {
                //        return false;
                //    }
                //    return true;
                //}

                // This was derived by trial and error, but doesn't seem quite right
                if (IsAutomaticLink)
                {
                    return false;
                }
                return true;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("[EXTERNALNAME]\n");
            sb.Append("    .options      = ").Append(field_1_option_flag).Append("\n");
            sb.Append("    .ix      = ").Append(field_2_ixals).Append("\n");
            sb.Append("    .name    = ").Append(field_4_name).Append("\n");
            if (field_5_name_definition != null)
            {
                Ptg[] ptgs = field_5_name_definition.Tokens;
                for (int i = 0; i < ptgs.Length; i++)
                {
                    Ptg ptg = ptgs[i];
                    sb.Append(ptg.ToString()).Append(ptg.RVAType).Append("\n");
                }
            }
            sb.Append("[/EXTERNALNAME]\n");
            return sb.ToString();
        }
    }
}