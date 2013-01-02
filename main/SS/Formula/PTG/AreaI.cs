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
namespace NPOI.SS.Formula.PTG
{
    using System;
    /**
     * Common interface for AreaPtg and Area3DPtg, and their
     *  child classes.
     */
    public interface AreaI
    {
        /**
         * @return the first row in the area
         */
        int FirstRow{get;}

        /**
         * @return last row in the range (x2 in x1,y1-x2,y2)
         */
        int LastRow { get; }

        /**
         * @return the first column number in the area.
         */
        int FirstColumn{get;}

        /**
         * @return lastcolumn in the area
         */
        int LastColumn{get;}
    }

    	public class OffsetArea: AreaI {

		private int _firstColumn;
		private int _firstRow;
		private int _lastColumn;
		private int _lastRow;

		public OffsetArea(int baseRow, int baseColumn, int relFirstRowIx, int relLastRowIx,
				int relFirstColIx, int relLastColIx) {
			_firstRow = baseRow + Math.Min(relFirstRowIx, relLastRowIx);
			_lastRow = baseRow + Math.Max(relFirstRowIx, relLastRowIx);
			_firstColumn = baseColumn + Math.Min(relFirstColIx, relLastColIx);
			_lastColumn = baseColumn + Math.Max(relFirstColIx, relLastColIx);
		}

		public int FirstColumn 
        {
            get
            {
                return _firstColumn;
            }
		}

		public int FirstRow
        {
            get
            {
                return _firstRow;
            }
		}

		public int LastColumn
        {
            get
            {
                return _lastColumn;
            }
		}

		public int LastRow
        {
            get
            {
                return _lastRow;
            }
		}
	}
}