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


    /**
     * Page Setup and page margins Settings for the worksheet.
     */
    public class XSSFPrintSetup : IPrintSetup
    {
        private CT_Worksheet ctWorksheet;
        private CT_PageSetup pageSetup;
        private CT_PageMargins pageMargins;

        public XSSFPrintSetup(CT_Worksheet worksheet)
        {
            this.ctWorksheet = worksheet;

            if (ctWorksheet.IsSetPageSetup())
            {
                this.pageSetup = ctWorksheet.pageSetup;
            }
            else
            {
                this.pageSetup = ctWorksheet.AddNewPageSetup();
            }
            if (ctWorksheet.IsSetPageMargins())
            {
                this.pageMargins = ctWorksheet.pageMargins;
            }
            else
            {
                this.pageMargins = ctWorksheet.AddNewPageMargins();
            }
        }


        /**
         * Set the paper size as enum value.
         *
         * @param size value for the paper size.
         */
        public void SetPaperSize(PaperSize size)
        {
            PaperSize = ((short)(size + 1));
        }








        /**
         * Orientation of the page: landscape - portrait.
         *
         * @return Orientation of the page
         * @see PrintOrientation
         */
        public PrintOrientation Orientation
        {
            get
            {
                ST_Orientation? val = pageSetup.orientation;
                return val == null ? PrintOrientation.DEFAULT : PrintOrientation.ValueOf((int)val);
            }
            set 
            {
                ST_Orientation v = (ST_Orientation)(value.Value);
                pageSetup.orientation = (v);
            }
        }


        public PrintCellComments GetCellComment()
        {
            ST_CellComments? val = pageSetup.cellComments;
            return val == null ? PrintCellComments.NONE : PrintCellComments.ValueOf((int)val);
        }


        /**
         * Get print page order.
         *
         * @return PageOrder
         */
        public PageOrder PageOrder
        {
            get
            {
                return PageOrder.ValueOf((int)pageSetup.pageOrder);
            }
            set 
            {
                ST_PageOrder v = (ST_PageOrder)value.Value;
                pageSetup.pageOrder = (v);
            }
        }

        /**
         * Returns the paper size.
         *
         * @return short - paper size
         */
        public short PaperSize
        {
            get
            {
                return (short)pageSetup.paperSize;
            }
            set 
            {
                pageSetup.paperSize = (uint)value;
            }
        }

        /**
         * Returns the paper size as enum.
         *
         * @return PaperSize paper size
         * @see PaperSize
         */
        public PaperSize GetPaperSizeEnum()
        {
            return (PaperSize)(PaperSize - 1);
        }

        /**
         * Returns the scale.
         *
         * @return short - scale
         */
        public short Scale
        {
            get
            {
                return (short)pageSetup.scale;
            }
            set 
            {
                if (value < 10 || value > 400) 
                    throw new POIXMLException("Scale value not accepted: you must choose a value between 10 and 400.");
                pageSetup.scale = (uint)value;
            }
        }

        /**
         * Set the page numbering start.
         * Page number for first printed page. If no value is specified, then 'automatic' is assumed.
         *
         * @return page number for first printed page
         */
        public short PageStart
        {
            get
            {
                return (short)pageSetup.firstPageNumber;
            }
            set 
            {
                pageSetup.firstPageNumber = (uint)value;
            }
        }

        /**
         * Returns the number of pages wide to fit sheet in.
         *
         * @return number of pages wide to fit sheet in
         */
        public short FitWidth
        {
            get
            {
                return (short)pageSetup.fitToWidth;
            }
            set 
            {
                pageSetup.fitToWidth = (uint)value;
            }
        }

        /**
         * Returns the number of pages high to fit the sheet in.
         *
         * @return number of pages high to fit the sheet in
         */
        public short FitHeight
        {
            get
            {
                return (short)pageSetup.fitToHeight;
            }
            set 
            {
                pageSetup.fitToHeight = (uint)value;
            }
        }

        /**
         * Returns the left to right print order.
         *
         * @return left to right print order
         */
        public bool LeftToRight
        {
            get
            {
                return PageOrder == PageOrder.OVER_THEN_DOWN;
            }
            set 
            {
                if (value)
                    PageOrder = (PageOrder.OVER_THEN_DOWN);
            }
        }

        /**
         * Returns the landscape mode.
         *
         * @return landscape mode
         */
        public bool Landscape
        {
            get
            {
                return Orientation == PrintOrientation.LANDSCAPE;
            }
            set 
            {
                if (value)
                    Orientation =(PrintOrientation.LANDSCAPE);
            }
        }

        /**
         * Use the printer's defaults Settings for page Setup values and don't use the default values
         * specified in the schema. For example, if dpi is not present or specified in the XML, the
         * application shall not assume 600dpi as specified in the schema as a default and instead
         * shall let the printer specify the default dpi.
         *
         * @return valid Settings
         */
        public bool ValidSettings
        {
            get
            {
                return pageSetup.usePrinterDefaults;
            }
            set 
            {
                pageSetup.usePrinterDefaults = value;
            }
        }

        /**
         * Returns the black and white Setting.
         *
         * @return black and white Setting
         */
        public bool NoColor
        {
            get
            {
                return pageSetup.blackAndWhite;
            }
            set 
            {
                pageSetup.blackAndWhite = value;
            }
        }

        /**
         * Returns the draft mode.
         *
         * @return draft mode
         */
        public bool Draft
        {
            get
            {
                return pageSetup.draft;
            }
            set 
            {
                pageSetup.draft = value;
            }
        }

        /**
         * Returns the print notes.
         *
         * @return print notes
         */
        public bool Notes
        {
            get
            {
                return GetCellComment() == PrintCellComments.AS_DISPLAYED;
            }
            set 
            {
                if (value)
                {
                    pageSetup.cellComments = (ST_CellComments.asDisplayed);
                }
            }
        }

        /**
         * Returns the no orientation.
         *
         * @return no orientation
         */
        public bool NoOrientation
        {
            get
            {
                return Orientation == PrintOrientation.DEFAULT;
            }
            set 
            {
                if (value)
                {
                    Orientation = (PrintOrientation.DEFAULT);
                }
            }
        }

        /**
         * Returns the use page numbers.
         *
         * @return use page numbers
         */
        public bool UsePage
        {
            get
            {
                return pageSetup.useFirstPageNumber;
            }
            set 
            {
                pageSetup.useFirstPageNumber = (value);
            }
        }

        /**
         * Returns the horizontal resolution.
         *
         * @return horizontal resolution
         */
        public short HResolution
        {
            get
            {
                return (short)pageSetup.horizontalDpi;
            }
            set 
            {
                pageSetup.horizontalDpi = (uint)value;
            }
        }

        /**
         * Returns the vertical resolution.
         *
         * @return vertical resolution
         */
        public short VResolution
        {
            get
            {
                return (short)pageSetup.verticalDpi;
            }
            set 
            {
                pageSetup.verticalDpi = (uint)value;
            }
        }

        /**
         * Returns the header margin.
         *
         * @return header margin
         */
        public double HeaderMargin
        {
            get
            {
                return pageMargins.header;
            }
            set 
            {
                pageMargins.header = (value);
            }
        }

        /**
         * Returns the footer margin.
         *
         * @return footer margin
         */
        public double FooterMargin
        {
            get
            {
                return pageMargins.footer;
            }
            set 
            {
                pageMargins.footer = value;
            }
        }

        /**
         * Returns the number of copies.
         *
         * @return number of copies
         */
        public short Copies
        {
            get
            {
                return (short)pageSetup.copies;
            }
            set 
            {
                pageSetup.copies = (uint)value;
            }
        }


        #region IPrintSetup Members

        public DisplayCellErrorType CellError
        {
            get
            {
                return (DisplayCellErrorType)pageSetup.errors;
            }
            set
            {
                pageSetup.errors = (ST_PrintError)value;
            }
        }

        public bool EndNote
        {
            get
            {
                
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        #endregion
    }
}



