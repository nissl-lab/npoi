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

namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.SS.UserModel;

    public class HSSFCreationHelper : CreationHelper
    {
        private HSSFWorkbook workbook;
        private HSSFDataFormat dataFormat;

        public HSSFCreationHelper(HSSFWorkbook wb)
        {
            workbook = wb;

            // Create the things we only ever need one of
            dataFormat = new HSSFDataFormat(workbook.Workbook);
        }

        public NPOI.SS.UserModel.RichTextString CreateRichTextString(String text)
        {
            return new HSSFRichTextString(text);
        }

        public NPOI.SS.UserModel.DataFormat CreateDataFormat()
        {
            return dataFormat;
        }

        public NPOI.SS.UserModel.Hyperlink CreateHyperlink(HyperlinkType type)
        {
            return new HSSFHyperlink(type);
        }

        /**
         * Creates a HSSFFormulaEvaluator, the object that Evaluates formula cells.
         *
         * @return a HSSFFormulaEvaluator instance
         */
        public NPOI.SS.UserModel.FormulaEvaluator CreateFormulaEvaluator()
        {
            return new HSSFFormulaEvaluator(workbook);
        }

        /**
         * Creates a HSSFClientAnchor. Use this object to position drawing object in a sheet
         *
         * @return a HSSFClientAnchor instance
         * @see NPOI.SS.usermodel.Drawing
         */
        public NPOI.SS.UserModel.ClientAnchor CreateClientAnchor()
        {
            return new HSSFClientAnchor();
        }
    }
}

