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
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
namespace NPOI.XSSF.UserModel
{
    public class XSSFCreationHelper : ICreationHelper
    {
        private XSSFWorkbook workbook;

        public XSSFCreationHelper(XSSFWorkbook wb)
        {
            workbook = wb;
        }

        /**
         * Creates a new XSSFRichTextString for you.
         */
        public IRichTextString CreateRichTextString(String text)
        {
            XSSFRichTextString rt = new XSSFRichTextString(text);
            rt.SetStylesTableReference(workbook.GetStylesSource());
            return rt;
        }

        public IDataFormat CreateDataFormat()
        {
            return workbook.CreateDataFormat();
        }

        public IHyperlink CreateHyperlink(HyperlinkType type)
        {
            return new XSSFHyperlink(type);
        }

        /**
         * Creates a XSSFFormulaEvaluator, the object that Evaluates formula cells.
         *
         * @return a XSSFFormulaEvaluator instance
         */
        public IFormulaEvaluator CreateFormulaEvaluator()
        {
            return new XSSFFormulaEvaluator(workbook);
        }

        /**
         * Creates a XSSFClientAnchor. Use this object to position Drawing object in
         * a sheet
         *
         * @return a XSSFClientAnchor instance
         * @see NPOI.ss.usermodel.Drawing
         */
        public IClientAnchor CreateClientAnchor()
        {
            return new XSSFClientAnchor();
        }
    }

}

