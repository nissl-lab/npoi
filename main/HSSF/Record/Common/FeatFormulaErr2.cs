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

namespace NPOI.HSSF.Record.Common
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * FeatFormulaErr2 (Formula Evaluation Shared Feature) common record part
     * 
     * This record part specifies Formula Evaluation &amp; Error Ignoring data 
     *  for a sheet, stored as part of a Shared Feature. It can be found in 
     *  records such as {@link FeatRecord}.
     * For the full meanings of the flags, see pages 669 and 670
     *  of the Excel binary file format documentation.
     */
    public class FeatFormulaErr2 : SharedFeature
    {
        static BitField checkCalculationErrors =
            BitFieldFactory.GetInstance(0x01);
        static BitField checkEmptyCellRef =
            BitFieldFactory.GetInstance(0x02);
        static BitField checkNumbersAsText =
            BitFieldFactory.GetInstance(0x04);
        static BitField checkInconsistentRanges =
            BitFieldFactory.GetInstance(0x08);
        static BitField checkInconsistentFormulas =
            BitFieldFactory.GetInstance(0x10);
        static BitField checkDateTimeFormats =
            BitFieldFactory.GetInstance(0x20);
        static BitField checkUnprotectedFormulas =
            BitFieldFactory.GetInstance(0x40);
        static BitField performDataValidation =
            BitFieldFactory.GetInstance(0x80);

        /**
         * What errors we should ignore
         */
        private int errorCheck;


        public FeatFormulaErr2() { }

        public FeatFormulaErr2(RecordInputStream in1)
        {
            errorCheck = in1.ReadInt();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append(" [FEATURE FORMULA ERRORS]\n");
            buffer.Append("  checkCalculationErrors    = ");
            buffer.Append("  checkEmptyCellRef         = ");
            buffer.Append("  checkNumbersAsText        = ");
            buffer.Append("  checkInconsistentRanges   = ");
            buffer.Append("  checkInconsistentFormulas = ");
            buffer.Append("  checkDateTimeFormats      = ");
            buffer.Append("  checkUnprotectedFormulas  = ");
            buffer.Append("  performDataValidation     = ");
            buffer.Append(" [/FEATURE FORMULA ERRORS]\n");
            return buffer.ToString();
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(errorCheck);
        }

        public int DataSize
        {
            get
            {
                return 4;
            }
        }

        public int RawErrorCheckValue
        {
            get
            {
                return errorCheck;
            }
        }

        public bool CheckCalculationErrors
        {
            get
            {
                return checkCalculationErrors.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.checkCalculationErrors.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckEmptyCellRef
        {
            get
            {
                return checkEmptyCellRef.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.checkEmptyCellRef.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckNumbersAsText
        {
            get
            {
                return checkNumbersAsText.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.checkNumbersAsText.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckInconsistentRanges
        {
            get
            {
                return checkInconsistentRanges.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.checkInconsistentRanges.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckInconsistentFormulas
        {
            get
            {
                return checkInconsistentFormulas.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.checkInconsistentFormulas.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckDateTimeFormats
        {
            get
            {
                return checkDateTimeFormats.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.checkDateTimeFormats.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckUnprotectedFormulas
        {
            get
            {
                return checkUnprotectedFormulas.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.checkUnprotectedFormulas.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool PerformDataValidation
        {
            get
            {
                return performDataValidation.IsSet(errorCheck);
            }
            set
            {
                FeatFormulaErr2.performDataValidation.SetBoolean(
                    errorCheck, value);
            }
        }
    }

}