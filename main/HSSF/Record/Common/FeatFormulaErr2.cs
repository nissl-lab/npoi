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
        private static BitField CHECK_CALCULATION_ERRORS = BitFieldFactory.GetInstance(0x01);
        private static BitField CHECK_EMPTY_CELL_REF = BitFieldFactory.GetInstance(0x02);
        private static BitField CHECK_NUMBERS_AS_TEXT = BitFieldFactory.GetInstance(0x04);
        private static BitField CHECK_INCONSISTENT_RANGES = BitFieldFactory.GetInstance(0x08);
        private static BitField CHECK_INCONSISTENT_FORMULAS = BitFieldFactory.GetInstance(0x10);
        private static BitField CHECK_DATETIME_FORMATS = BitFieldFactory.GetInstance(0x20);
        private static BitField CHECK_UNPROTECTED_FORMULAS = BitFieldFactory.GetInstance(0x40);
        private static BitField PERFORM_DATA_VALIDATION = BitFieldFactory.GetInstance(0x80);

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
                return CHECK_CALCULATION_ERRORS.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.CHECK_CALCULATION_ERRORS.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckEmptyCellRef
        {
            get
            {
                return CHECK_EMPTY_CELL_REF.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.CHECK_EMPTY_CELL_REF.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckNumbersAsText
        {
            get
            {
                return CHECK_NUMBERS_AS_TEXT.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.CHECK_NUMBERS_AS_TEXT.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckInconsistentRanges
        {
            get
            {
                return CHECK_INCONSISTENT_RANGES.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.CHECK_INCONSISTENT_RANGES.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckInconsistentFormulas
        {
            get
            {
                return CHECK_INCONSISTENT_FORMULAS.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.CHECK_INCONSISTENT_FORMULAS.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckDateTimeFormats
        {
            get
            {
                return CHECK_DATETIME_FORMATS.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.CHECK_DATETIME_FORMATS.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool CheckUnprotectedFormulas
        {
            get
            {
                return CHECK_UNPROTECTED_FORMULAS.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.CHECK_UNPROTECTED_FORMULAS.SetBoolean(
                        errorCheck, value);
            }
        }

        public bool PerformDataValidation
        {
            get
            {
                return PERFORM_DATA_VALIDATION.IsSet(errorCheck);
            }
            set
            {
                errorCheck = FeatFormulaErr2.PERFORM_DATA_VALIDATION.SetBoolean(
                    errorCheck, value);
            }
        }
    }

}