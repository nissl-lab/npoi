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

using System;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;

namespace NPOI.XSSF.UserModel
{
    /**
     * A client anchor is attached to an excel worksheet.  It anchors against
     * top-left and bottom-right cells.
     *
     * @author Yegor Kozlov
     */
    public class XSSFClientAnchor : XSSFAnchor, IClientAnchor
    {
        private int anchorType;

        /**
         * Starting anchor point
         */
        private CT_Marker cell1;

        /**
         * Ending anchor point
         */
        private CT_Marker cell2;

        /**
         * Creates a new client anchor and defaults all the anchor positions to 0.
         */
        public XSSFClientAnchor()
        {
            cell1 = new CT_Marker();
            cell1.col= (0);
            cell1.colOff=(0);
            cell1.row=(0);
            cell1.rowOff=(0);
            cell2 = new CT_Marker();
            cell2.col=(0);
            cell2.colOff=(0);
            cell2.row=(0);
            cell2.rowOff=(0);
        }

        /**
         * Creates a new client anchor and Sets the top-left and bottom-right
         * coordinates of the anchor.
         *
         * @param dx1  the x coordinate within the first cell.
         * @param dy1  the y coordinate within the first cell.
         * @param dx2  the x coordinate within the second cell.
         * @param dy2  the y coordinate within the second cell.
         * @param col1 the column (0 based) of the first cell.
         * @param row1 the row (0 based) of the first cell.
         * @param col2 the column (0 based) of the second cell.
         * @param row2 the row (0 based) of the second cell.
         */
        public XSSFClientAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
            : this()
        {

            cell1.col = (col1);
            cell1.colOff = (dx1);
            cell1.row = (row1);
            cell1.rowOff = (dy1);
            cell2.col = (col2);
            cell2.colOff = (dx2);
            cell2.row = (row2);
            cell2.rowOff = (dy2);
        }

        /**
         * Create XSSFClientAnchor from existing xml beans
         *
         * @param cell1 starting anchor point
         * @param cell2 ending anchor point
         */
        internal XSSFClientAnchor(CT_Marker cell1, CT_Marker cell2)
        {
            this.cell1 = cell1;
            this.cell2 = cell2;
        }

 



        public override bool Equals(Object o)
        {
            if (o == null || !(o is XSSFClientAnchor)) return false;

            XSSFClientAnchor anchor = (XSSFClientAnchor)o;
            return Dx1 == anchor.Dx1 &&
                Dx2 == anchor.Dx2 &&
                Dy1 == anchor.Dy1 &&
                Dy2 == anchor.Dy2 &&
                Col1 == anchor.Col1 &&
                Col2 == anchor.Col2 &&
                Row1 == anchor.Row1 &&
                Row2 == anchor.Row2;

        }


        public override String ToString()
        {
            return "from : " + cell1.ToString() + "; to: " + cell2.ToString();
        }

        /**
         * Return starting anchor point
         *
         * @return starting anchor point
         */

        internal CT_Marker From
        {
            get
            {
                return cell1;
            }
            set 
            {
                cell1 = value;
            }
        }

        /**
         * Return ending anchor point
         *
         * @return ending anchor point
         */

        internal CT_Marker To
        {
            get
            {
                return cell2;
            }
            set 
            {
                cell2 = value;
            }
        }


        internal bool IsSet()
        {
            return !(cell1.col == 0 && cell2.col == 0 &&
                     cell1.row == 0 && cell2.row == 0);
        }

        #region IClientAnchor Members
        public override int Dx1
        {
            get
            {
                return (int)cell1.colOff;
            }
            set
            {
                cell1.colOff = value;
            }
        }

        public override int Dy1
        {
            get
            {
                return (int)cell1.rowOff;
            }
            set
            {
                cell1.rowOff = value;
            }
        }

        public override int Dy2
        {
            get
            {
                return (int)cell2.rowOff;
            }
            set
            {
                cell2.rowOff = value;
            }
        }

        public override int Dx2
        {
            get
            {
                return (int)cell2.colOff;
            }
            set
            {
                cell2.colOff = value;
            }
        }
        public AnchorType AnchorType
        {
            get
            {
                return (AnchorType)this.anchorType;
            }
            set
            {
                this.anchorType = (int)value;
            }
        }

        public int Col1
        {
            get
            {
                return cell1.col;
            }
            set
            {
                cell1.col=value;
            }
        }

        public int Col2
        {
            get
            {
                return cell2.col;
            }
            set
            {
                cell2.col = value;
            }
        }

        public int Row1
        {
            get
            {
                return cell1.row;
            }
            set
            {
                cell1.row = value;
            }
        }

        public int Row2
        {
            get
            {
                return cell2.row;
            }
            set
            {
                cell2.row = value;
            }
        }

        #endregion
    }
}



