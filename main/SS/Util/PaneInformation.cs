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

namespace NPOI.SS.Util
{

    /// <summary>
    /// Holds information regarding a split plane or freeze plane for a sheet.
    /// </summary>
    public class PaneInformation
    {
        /// <summary>
        /// Constant for active pane being the lower right*/
        /// </summary>
        public const byte PANE_LOWER_RIGHT = (byte)0;
        /// <summary>
        /// Constant for active pane being the upper right*/
        /// </summary>
        public const byte PANE_UPPER_RIGHT = (byte)1;
        /// <summary>
        /// Constant for active pane being the lower left*/
        /// </summary>
        public const byte PANE_LOWER_LEFT = (byte)2;
        /// <summary>
        /// Constant for active pane being the upper left*/
        /// </summary>
        public const byte PANE_UPPER_LEFT = (byte)3;

        private short x;
        private short y;
        private short topRow;
        private short leftColumn;
        private byte activePane;
        private bool frozen = false;

        public PaneInformation(short x, short y, short top, short left, byte active, bool frozen)
        {
            this.x = x;
            this.y = y;
            this.topRow = top;
            this.leftColumn = left;
            this.activePane = active;
            this.frozen = frozen;
        }


        /// <summary>
        /// Returns the vertical position of the split.
        /// </summary>
        /// <return>0 if there is no vertical spilt,
        /// or for a freeze pane the number of columns in the TOP pane,
        /// or for a split plane the position of the split in 1/20th of a point.
        /// </return>
        public short VerticalSplitPosition
        {
            get { return x; }
        }

        /// <summary>
        /// Returns the horizontal position of the split.
        /// </summary>
        /// <return>0 if there is no horizontal spilt,
        /// or for a freeze pane the number of rows in the LEFT pane,
        /// or for a split plane the position of the split in 1/20th of a point.
        /// </return>
        public short HorizontalSplitPosition
        {
            get { return y; }
        }

        /// <summary>
        /// For a horizontal split returns the top row in the BOTTOM pane.
        /// </summary>
        /// <return>0 if there is no horizontal split, or the top row of the bottom pane.</return>
        public short HorizontalSplitTopRow
        {
            get { return topRow; }
        }

        /// <summary>
        /// For a vertical split returns the left column in the RIGHT pane.
        /// </summary>
        /// <return>0 if there is no vertical split, or the left column in the RIGHT pane.</return>
        public short VerticalSplitLeftColumn
        {
            get { return leftColumn; }
        }

        /// <summary>
        /// Returns the active pane
        /// </summary>
        /// <see cref="PANE_LOWER_RIGHT" />
        /// <see cref="PANE_UPPER_RIGHT" />
        /// <see cref="PANE_LOWER_LEFT" />
        /// <see cref="PANE_UPPER_LEFT" />
        /// <return>the active pane.</return>
        public byte ActivePane
        {
            get { return activePane; }
        }

        /// <summary>
        /// Returns true if this is a Freeze pane, false if it is a split pane.
        /// </summary>
        public bool IsFreezePane()
        {
            return frozen;
        }
    }
}