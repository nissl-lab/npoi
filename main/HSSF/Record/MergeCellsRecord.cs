
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


namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using System.Collections;

    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * Title: Merged Cells Record
     * 
     * Description:  Optional record defining a square area of cells to "merged" into
     *               one cell. 
     * REFERENCE:  NONE (UNDOCUMENTED PRESENTLY) 
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */
    public class MergeCellsRecord : StandardRecord, ICloneable
    {
        public const short sid = 0xe5;

        /** sometimes the regions array is shared with other MergedCellsRecords */
        private CellRangeAddress[] _regions;
        private int _startIndex;
        private int _numberOfRegions;


        public MergeCellsRecord(CellRangeAddress[] regions, int startIndex, int numberOfRegions)
        {
            _regions = regions;
            _startIndex = startIndex;
            _numberOfRegions = numberOfRegions;
        }

        /**
         * Constructs a MergedCellsRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public MergeCellsRecord(RecordInputStream in1)
        {
            int nRegions = in1.ReadUShort();
    	    CellRangeAddress[] cras = new CellRangeAddress[nRegions];
    	    for (int i = 0; i < nRegions; i++) 
            {
			    cras[i] = new CellRangeAddress(in1);
		    }
    	    _numberOfRegions = nRegions;
    	    _startIndex = 0;
    	    _regions = cras;
        }

        public IEnumerator GetEnumerator()
        {
            return _regions.GetEnumerator();
        }

        /**
         * Get the number of merged areas.  If this drops down to 0 you should just go
         * ahead and delete the record.
         * @return number of areas
         */

        public short NumAreas
        {
            get
            {
                return (short)_numberOfRegions;
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        /**
         * @return MergedRegion at the given index representing the area that is Merged (r1,c1 - r2,c2)
         */
        public CellRangeAddress GetAreaAt(int index)
        {
            return _regions[_startIndex + index];
        }
        protected override int DataSize
        {
            get
            {
                return CellRangeAddressList.GetEncodedSize(_numberOfRegions);
            }
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override void Serialize(ILittleEndianOutput out1)
        {
            int nItems = _numberOfRegions;
            out1.WriteShort(nItems);
            for (int i = 0; i < _numberOfRegions; i++)
            {
                _regions[_startIndex + i].Serialize(out1);
            }
        }
        

        public override String ToString()
        {
            StringBuilder retval = new StringBuilder();

            retval.Append("[MERGEDCELLS]").Append("\n");
            retval.Append("     .numregions =").Append(NumAreas)
                .Append("\n");
            for (int k = 0; k < _numberOfRegions; k++)
            {
                CellRangeAddress region = _regions[_startIndex + k];

                retval.Append("     .rowfrom    =").Append(region.FirstRow)
                    .Append("\n");
                retval.Append("     .rowto      =").Append(region.LastRow)
                    .Append("\n");
                retval.Append("     .colfrom    =").Append(region.FirstColumn)
                    .Append("\n");
                retval.Append("     .colto      =").Append(region.LastColumn)
                    .Append("\n");
            }
            retval.Append("[MERGEDCELLS]").Append("\n");
            return retval.ToString();
        }

        public override Object Clone()
        {
            int nRegions = _numberOfRegions;
            CellRangeAddress[] clonedRegions = new CellRangeAddress[nRegions];
            for (int i = 0; i < clonedRegions.Length; i++)
            {
                clonedRegions[i] = _regions[_startIndex + i].Copy();
            }
            return new MergeCellsRecord(clonedRegions, 0, nRegions);
        }
    }
}