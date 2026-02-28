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


namespace NPOI.HSSF.UserModel
{
    using NPOI.HSSF.Record;

    /// <summary>
    /// Used to modify the print Setup.
    /// @author Shawn Laubach (slaubach at apache dot org)
    /// </summary>
    public class HSSFPrintSetup :NPOI.SS.UserModel.IPrintSetup
    {
        PrintSetupRecord printSetupRecord;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFPrintSetup"/> class.
        /// </summary>
        /// <param name="printSetupRecord">Takes the low level print Setup record.</param>
        public HSSFPrintSetup(PrintSetupRecord printSetupRecord)
        {
            this.printSetupRecord = printSetupRecord;
        }

        /// <summary>
        /// Gets or sets the size of the paper.
        /// </summary>
        /// <value>The size of the paper.</value>
        public short PaperSize
        {
            get
            {
                return printSetupRecord.PaperSize;
            }
            set 
            {
                printSetupRecord.PaperSize = value;
            }
        }

        /// <summary>
        /// Gets or sets the scale.
        /// </summary>
        /// <value>The scale.</value>
        public short Scale
        {
            get
            {
                return printSetupRecord.Scale;
            }
            set 
            {
                printSetupRecord.Scale = value;
            }
        }

        /// <summary>
        /// Gets or sets the page start.
        /// </summary>
        /// <value>The page start.</value>
        public short PageStart
        {
            get
            {
                return printSetupRecord.PageStart;
            }
            set 
            {
                printSetupRecord.PageStart=value;
            }
        }

        /// <summary>
        /// Gets or sets the number of pages wide to fit sheet in.
        /// </summary>
        /// <value>the number of pages wide to fit sheet in</value>
        public short FitWidth
        {
            get
            {
                return printSetupRecord.FitWidth;
            }
            set 
            {
                printSetupRecord.FitWidth = value;
            }
        }

        /// <summary>
        /// Gets or sets number of pages high to fit the sheet in
        /// </summary>
        /// <value>number of pages high to fit the sheet in.</value>
        public short FitHeight
        {
            get
            {
                return printSetupRecord.FitHeight;
            }
            set 
            {
                printSetupRecord.FitHeight = value;
            }
        }

        /// <summary>
        /// Gets or sets the bit flags for the options.
        /// </summary>
        /// <value>the bit flags for the options.</value>
        public short Options
        {
            get { return printSetupRecord.Options; }
            set { printSetupRecord.Options=(value); }
        }

        /// <summary>
        /// Gets or sets the left to right print order.
        /// </summary>
        /// <value>the left to right print order.</value>
        public bool LeftToRight
        {
            get
            {
                return printSetupRecord.LeftToRight;
            }
            set 
            {
                printSetupRecord.LeftToRight = value;
            }
        }

        /// <summary>
        /// Gets or sets the landscape mode.
        /// </summary>
        /// <value>the landscape mode.</value>
        public bool Landscape
        {
            get
            {
                return !printSetupRecord.Landscape;
            }
            set 
            {
                printSetupRecord.Landscape = !value;
            }
        }

        /// <summary>
        /// Gets or sets the valid Settings.
        /// </summary>
        /// <value>the valid Settings.</value>
        public bool ValidSettings
        {
            get
            {
                return printSetupRecord.ValidSettings;
            }
            set 
            {
                printSetupRecord.ValidSettings = value;
            }
        }

        /// <summary>
        /// Gets or sets the black and white Setting.
        /// </summary>
        /// <value>black and white Setting</value>
        public bool NoColor
        {
            get
            {
                return printSetupRecord.NoColor;
            }
            set 
            {
                printSetupRecord.NoColor=value;
            }
        }

        public bool EndNote
        {
            get { return printSetupRecord.EndNote; }
            set 
            {
                printSetupRecord.EndNote = value;
            }
        }
        public NPOI.SS.UserModel.DisplayCellErrorType CellError
        {
            get { return (NPOI.SS.UserModel.DisplayCellErrorType)printSetupRecord.CellError; }
            set { printSetupRecord.CellError=(short)value; }
        }

        /// <summary>
        /// Gets or sets the draft mode.
        /// </summary>
        /// <value>the draft mode.</value>
        public bool Draft
        {
            get
            {
                return printSetupRecord.Draft;
            }
            set 
            {
                printSetupRecord.Draft = value;
            }
        }

        /// <summary>
        /// Gets or sets the print notes.
        /// </summary>
        /// <value>the print notes.</value>
        public bool Notes
        {
            get
            {
                return printSetupRecord.Notes;
            }
            set 
            {
                printSetupRecord.Notes = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [no orientation].
        /// </summary>
        /// <value><c>true</c> if [no orientation]; otherwise, <c>false</c>.</value>
        public bool NoOrientation
        {
            get
            {
                return printSetupRecord.NoOrientation;
            }
            set 
            {
                printSetupRecord.NoOrientation = value;
            }
        }

        /// <summary>
        /// Gets or sets the use page numbers.  
        /// </summary>
        /// <value>use page numbers.  </value>
        public bool UsePage
        {
            get
            {
                return printSetupRecord.UsePage;
            }
            set 
            {
                printSetupRecord.UsePage = value;
            }
        }

        /// <summary>
        /// Gets or sets the horizontal resolution.
        /// </summary>
        /// <value>the horizontal resolution.</value>
        public short HResolution
        {
            get
            {
                return printSetupRecord.HResolution;
            }
            set
            {
                printSetupRecord.HResolution = value;
            }
        }

        /// <summary>
        /// Gets or sets the vertical resolution.
        /// </summary>
        /// <value>the vertical resolution.</value>
        public short VResolution
        {
            get
            {
                return printSetupRecord.VResolution;
            }
            set 
            {
                printSetupRecord.VResolution = value;
            }
        }

        /// <summary>
        /// Gets or sets the header margin.
        /// </summary>
        /// <value>The header margin.</value>
        public double HeaderMargin
        {
            get
            {
                return printSetupRecord.HeaderMargin;
            }
            set
            {
                printSetupRecord.HeaderMargin=value;
            }
        }

        /// <summary>
        /// Gets or sets the footer margin.
        /// </summary>
        /// <value>The footer margin.</value>
        public double FooterMargin
        {
            get
            {
                return printSetupRecord.FooterMargin;
            }
            set 
            {
                printSetupRecord.FooterMargin = value;
            }
        }

        /// <summary>
        /// Gets or sets the number of copies.
        /// </summary>
        /// <value>the number of copies.</value>
        public short Copies
        {
            get
            {
                return printSetupRecord.Copies;
            }
            set 
            {
                printSetupRecord.Copies = value;
            }
        }
    }
}