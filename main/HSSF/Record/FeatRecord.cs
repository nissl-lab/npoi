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

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record.Common;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * Title: Feat (Feature) Record
     * 
     * This record specifies Shared Features data. It is normally paired
     *  up with a {@link FeatHdrRecord}.
     */
    public class FeatRecord : StandardRecord
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(FeatRecord));
        public const short sid = 0x0868;

        private FtrHeader futureHeader;

        /**
         * See SHAREDFEATURES_* on {@link FeatHdrRecord}
         */
        private int isf_sharedFeatureType;
        private byte reserved1; // Should always be zero
        private long reserved2; // Should always be zero
        /** Only matters if type is ISFFEC2 */
        private long cbFeatData;
        private int reserved3; // Should always be zero
        private CellRangeAddress[] cellRefs;

        /**
         * Contents depends on isf_sharedFeatureType :
         *  ISFPROTECTION -> FeatProtection 
         *  ISFFEC2       -> FeatFormulaErr2
         *  ISFFACTOID    -> FeatSmartTag
         */
        private SharedFeature sharedFeature;

        public FeatRecord()
        {
            futureHeader = new FtrHeader();
            futureHeader.RecordType = (/*setter*/sid);
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public FeatRecord(RecordInputStream in1)
        {
            futureHeader = new FtrHeader(in1);

            isf_sharedFeatureType = in1.ReadShort();
            reserved1 = (byte)in1.ReadByte();
            reserved2 = in1.ReadInt();
            int cref = in1.ReadUShort();
            cbFeatData = in1.ReadInt();
            reserved3 = in1.ReadShort();

            cellRefs = new CellRangeAddress[cref];
            for (int i = 0; i < cellRefs.Length; i++)
            {
                cellRefs[i] = new CellRangeAddress(in1);
            }

            switch (isf_sharedFeatureType)
            {
                case FeatHdrRecord.SHAREDFEATURES_ISFPROTECTION:
                    sharedFeature = new FeatProtection(in1);
                    break;
                case FeatHdrRecord.SHAREDFEATURES_ISFFEC2:
                    sharedFeature = new FeatFormulaErr2(in1);
                    break;
                case FeatHdrRecord.SHAREDFEATURES_ISFFACTOID:
                    sharedFeature = new FeatSmartTag(in1);
                    break;
                default:
                    logger.Log(POILogger.ERROR, "Unknown Shared Feature " + isf_sharedFeatureType + " found!");
                    break;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[SHARED FEATURE]\n");

            // TODO ...

            buffer.Append("[/SHARED FEATURE]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            futureHeader.Serialize(out1);

            out1.WriteShort(isf_sharedFeatureType);
            out1.WriteByte(reserved1);
            out1.WriteInt((int)reserved2);
            out1.WriteShort(cellRefs.Length);
            out1.WriteInt((int)cbFeatData);
            out1.WriteShort(reserved3);

            for (int i = 0; i < cellRefs.Length; i++)
            {
                cellRefs[i].Serialize(out1);
            }

            sharedFeature.Serialize(out1);
        }

        protected override int DataSize
        {
            get
            {
                return 12 + 2 + 1 + 4 + 2 + 4 + 2 +
                    (cellRefs.Length * CellRangeAddress.ENCODED_SIZE)
                    + sharedFeature.DataSize;
            }
        }

        public int Isf_sharedFeatureType
        {
            get
            {
                return isf_sharedFeatureType;
            }
        }

        public long CbFeatData
        {
            get
            {
                return cbFeatData;
            }
            set
            {
                this.cbFeatData = value;
            }
        }


        public CellRangeAddress[] CellRefs
        {
            get
            {
                return cellRefs;
            }
            set
            {
                this.cellRefs = value;
            }
        }
        

        public SharedFeature SharedFeature
        {
            get
            {
                return sharedFeature;
            }
            set
            {
                this.sharedFeature = value;

                if (value is FeatProtection)
                {
                    isf_sharedFeatureType = FeatHdrRecord.SHAREDFEATURES_ISFPROTECTION;
                }
                if (value is FeatFormulaErr2)
                {
                    isf_sharedFeatureType = FeatHdrRecord.SHAREDFEATURES_ISFFEC2;
                }
                if (value is FeatSmartTag)
                {
                    isf_sharedFeatureType = FeatHdrRecord.SHAREDFEATURES_ISFFACTOID;
                }

                if (isf_sharedFeatureType == FeatHdrRecord.SHAREDFEATURES_ISFFEC2)
                {
                    cbFeatData = sharedFeature.DataSize;
                }
                else
                {
                    cbFeatData = 0;
                }
            }
        }
        

        //HACK: do a "cheat" Clone, see Record.java for more information
        public override Object Clone()
        {
            return CloneViaReserialise();
        }


    }
}
