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
using NPOI.OpenXmlFormats.Spreadsheet;
namespace NPOI.XSSF.UserModel
{

    //YK: TODO: this is only a prototype
    public class XSSFDialogsheet : XSSFSheet, ISheet
    {
        protected CT_Dialogsheet dialogsheet;

        protected XSSFDialogsheet(XSSFSheet sheet)
            : base(sheet.GetPackagePart(), sheet.GetPackageRelationship())
        {

            this.dialogsheet = new CT_Dialogsheet();
            this.worksheet = new CT_Worksheet();
        }

        public override IRow CreateRow(int rowNum)
        {
            return null;
        }

        protected CT_HeaderFooter GetSheetTypeHeaderFooter()
        {
            if (dialogsheet.headerFooter == null)
            {
                dialogsheet.headerFooter = (new CT_HeaderFooter());
            }
            return dialogsheet.headerFooter;
        }

        protected CT_SheetPr GetSheetTypeSheetPr()
        {
            if (dialogsheet.sheetPr == null)
            {
                dialogsheet.sheetPr = (new CT_SheetPr());
            }
            return dialogsheet.sheetPr;
        }

        protected CT_PageBreak GetSheetTypeColumnBreaks()
        {
            return null;
        }

        protected CT_SheetFormatPr GetSheetTypeSheetFormatPr()
        {
            if (dialogsheet.sheetFormatPr == null)
            {
                dialogsheet.sheetFormatPr = (new CT_SheetFormatPr());
            }
            return dialogsheet.sheetFormatPr;
        }

        protected CT_PageMargins GetSheetTypePageMargins()
        {
            if (dialogsheet.pageMargins == null)
            {
                dialogsheet.pageMargins = (new CT_PageMargins());
            }
            return dialogsheet.pageMargins;
        }

        protected CT_PageBreak GetSheetTypeRowBreaks()
        {
            return null;
        }

        protected CT_SheetViews GetSheetTypeSheetViews()
        {
            if (dialogsheet.sheetViews == null)
            {
                dialogsheet.sheetViews = (new CT_SheetViews());
                dialogsheet.sheetViews.AddNewSheetView();
            }
            return dialogsheet.sheetViews;
        }

        protected CT_PrintOptions GetSheetTypePrintOptions()
        {
            if (dialogsheet.printOptions == null)
            {
                dialogsheet.printOptions = (new CT_PrintOptions());
            }
            return dialogsheet.printOptions;
        }

        protected CT_SheetProtection GetSheetTypeProtection()
        {
            if (dialogsheet.sheetProtection == null)
            {
                dialogsheet.sheetProtection = (new CT_SheetProtection());
            }
            return dialogsheet.sheetProtection;
        }

        public bool GetDialog()
        {
            return true;
        }
    }
}


