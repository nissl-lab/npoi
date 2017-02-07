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
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFCreationHelper : ICreationHelper
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(SXSSFCreationHelper));

        private SXSSFWorkbook wb;
        private XSSFCreationHelper helper;

        public SXSSFCreationHelper(SXSSFWorkbook workbook)
        {
            this.helper = new XSSFCreationHelper(workbook.XssfWorkbook);
            this.wb = workbook;
        }

        public IClientAnchor CreateClientAnchor()
        {
            return helper.CreateClientAnchor();
        }

        public IDataFormat CreateDataFormat()
        {
            return helper.CreateDataFormat();
        }

        public IFormulaEvaluator CreateFormulaEvaluator()
        {
            return new SXSSFFormulaEvaluator(wb);
        }

        public IHyperlink CreateHyperlink(HyperlinkType type)
        {
            return helper.CreateHyperlink(type);
        }

        public IRichTextString CreateRichTextString(string text)
        {
            logger.Log(POILogger.INFO, "SXSSF doesn't support Rich Text Strings, any formatting information will be lost");
            return new XSSFRichTextString(text);
        }

        //TODO: missing methods CreateExtendedColor()
    }
}
