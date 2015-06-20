
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
    using NPOI.Util;


    /**
     * Title:        Unknown Record (for debugging)
     * Description:  Unknown record just tells you the sid so you can figure out
     *               what records you are missing.  Also helps us Read/modify sheets we
     *               don't know all the records to.  (HSSF leaves these alone!) 
     * Company:      SuperLink Software, Inc.
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @author Glen Stampoultzis (glens at apache.org)
     */

    public class UnknownRecord : StandardRecord
    {
        //public const int PRINTSIZE_0033       = 0x0033;
        /*
     * Some Record IDs used by POI as 'milestones' in the record stream
     */
        public const int PLS_004D = 0x004D;
        public const int SHEETPR_0081 = 0x0081;
        public const int SORT_0090            = 0x0090;
        public const int STANDARDWIDTH_0099 = 0x0099;
        //public const int SCL_00A0 = 0x00A0;
        public const int BITMAP_00E9 = 0x00E9;
        public const int PHONETICPR_00EF = 0x00EF;
        public const int LABELRANGES_015F = 0x015F;
      	//public const int USERSVIEWBEGIN_01AA  = 0x01AA;
    	//public const int USERSVIEWEND_01AB    = 0x01AB;
        public const int QUICKTIP_0800 = 0x0800;
        //public const int SHEETEXT_0862 = 0x0862; // OOO calls this SHEETLAYOUT
        public const int SHEETPROTECTION_0867 = 0x0867;
        //public const int RANGEPROTECTION_0868 = 0x0868;
        public const int HEADER_FOOTER_089C   = 0x089C;
        public const int CODENAME_1BA = 0x01BA;
        public const int PLV_MAC = 0x08C8;
        private int _sid = 0;
        private byte[] _rawData = null;

        //public UnknownRecord()
        //{
        //}

        /**
         * @param id    id of the record -not Validated, just stored for serialization
         * @param data  the data
         */
        public UnknownRecord(int id, byte[] data)
        {
            _sid = id & 0xFFFF;
            _rawData = data;
        }


        /**
         * construct an Unknown record.  No fields are interperated and the record will
         * be Serialized in its original form more or less
         * @param in the RecordInputstream to Read the record from
         */

        public UnknownRecord(RecordInputStream in1)
        {
            _sid = in1.Sid;
            _rawData = in1.ReadRemainder();

            //if (false && GetBiffName(_sid) == null)
            //{
            //    // unknown sids in the range 0x0004-0x0013 are probably 'sub-records' of ObjectRecord
            //    // those sids are in a different number space.
            //    // TODO - put unknown OBJ sub-records in a different class
            //    System.Console.WriteLine("Unknown record 0x" + _sid.ToString("X"));
            //}
        }

	/**
	 * spit the record out AS IS. no interpretation or identification
	 */
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.Write(_rawData);
        }
        protected override int DataSize
        {
            get
            {
                return _rawData.Length;
            }
        }

        /**
         * print a sort of string representation ([UNKNOWN RECORD] id = x [/UNKNOWN RECORD])
         */

        public override String ToString()
        {
            String biffName = GetBiffName(_sid);
            if (biffName == null)
            {
                biffName = "UNKNOWNRECORD";
            }
            StringBuilder sb = new StringBuilder();

            sb.Append("[").Append(biffName).Append("] (0x");
            sb.Append(StringUtil.ToHexString(_sid).ToUpper() + ")\n");
            if (_rawData.Length > 0)
            {
                sb.Append("  rawData=").Append(HexDump.ToHex(_rawData)).Append("\n");
            }
            sb.Append("[/").Append(biffName).Append("]\n");
            return sb.ToString();

        }

        /**
 * These BIFF record types are known but still uninterpreted by POI
 *
 * @return the documented name of this BIFF record type, <code>null</code> if unknown to POI
 */
        public static String GetBiffName(int sid)
        {
            // Note to POI developers:
            // Make sure you delete the corresponding entry from
            // this method any time a new Record subclass is created.
            switch (sid)
            {
                //case PRINTSIZE_0033: return "PRINTSIZE";
                case PLS_004D: return "PLS";
                case 0x0050: return "DCON"; // Data Consolidation Information
                case 0x007F: return "IMDATA";
                case SHEETPR_0081: return "SHEETPR";
                case SORT_0090: return "SORT"; // Sorting Options
                case 0x0094: return "LHRECORD"; // .WK? File Conversion Information
                case STANDARDWIDTH_0099: return "STANDARDWIDTH"; //Standard Column Width
                case 0x009D: return "AUTOFILTERINFO"; // Drop-Down Arrow Count
                //case SCL_00A0: return "SCL"; // Window Zoom Magnification
                case 0x00AE: return "SCENMAN"; // Scenario Output Data

                case 0x00B2: return "SXVI";        // (pivot table) View Item
                case 0x00B4: return "SXIVD";       // (pivot table) Row/Column Field IDs
                case 0x00B5: return "SXLI";        // (pivot table) Line Item Array

                case 0x00D3: return "OBPROJ";
                case 0x00DC: return "PARAMQRY";
                case 0x00DE: return "OLESIZE";
                case BITMAP_00E9: return "BITMAP";
                case PHONETICPR_00EF: return "PHONETICPR";
                case 0x00F1: return "SXEX";        // PivotTable View Extended Information

                case LABELRANGES_015F: return "LABELRANGES";
                case 0x01BA: return "CODENAME";
                case 0x01A9: return "USERBVIEW";
                case 0x01AD: return "QSI";

                case 0x01C0: return "EXCEL9FILE";

                case 0x0802: return "QSISXTAG";   // Pivot Table and Query Table Extensions
                case 0x0803: return "DBQUERYEXT";
                case 0x0805: return "TXTQUERY";
                case 0x0810: return "SXVIEWEX9";  // Pivot Table Extensions

                case 0x0812: return "CONTINUEFRT";
                case QUICKTIP_0800: return "QUICKTIP";
                //case SHEETEXT_0862: return "SHEETEXT";
                case 0x0863: return "BOOKEXT";
                case 0x0864: return "SXADDL";    // Pivot Table Additional Info
                case SHEETPROTECTION_0867: return "SHEETPROTECTION";
                //case RANGEPROTECTION_0868: return "RANGEPROTECTION";
                case 0x086B: return "DATALABEXTCONTENTS";
                case 0x086C: return "CELLWATCH";
                case 0x0874: return "DROPDOWNOBJIDS";
                case 0x0876: return "DCONN";
                case 0x087B: return "CFEX";
                case 0x087C: return "XFCRC";
                case 0x087D: return "XFEXT";
                case 0x087F: return "CONTINUEFRT12";
                case 0x088B: return "PLV";
                case 0x088C: return "COMPAT12";
                case 0x088D: return "DXF";
                case 0x0892: return "STYLEEXT";
                case 0x0896: return "THEME";
                case 0x0897: return "GUIDTYPELIB";
                case 0x089A: return "MTRSETTINGS";
                case 0x089B: return "COMPRESSPICTURES";
                case HEADER_FOOTER_089C: return "HEADERFOOTER";
                case 0x08A1: return "SHAPEPROPSSTREAM";
                case 0x08A3: return "FORCEFULLCALCULATION";
                case 0x08A4: return "SHAPEPROPSSTREAM";
                case 0x08A5: return "TEXTPROPSSTREAM";
                case 0x08A6: return "RICHTEXTSTREAM";

                case 0x08C8: return "PLV{Mac Excel}";

            }
            if (IsObservedButUnknown(sid))
            {
                return "UNKNOWN-" + StringUtil.ToHexString(sid).ToUpper();
            }
            else
            {
                return "BIFF Record-" + StringUtil.ToHexString(sid).ToUpper();
            }

            //return null;
        }
        /**
 * @return <c>true</c> if the unknown record id has been observed in POI unit tests
 */
        private static bool IsObservedButUnknown(int sid)
        {
            switch (sid)
            {
                case 0x0033:
                // contains 2 bytes of data: 0x0001 or 0x0003
                case 0x0034:
                // Seems to be written by MSAccess
                // contains text "[Microsoft JET Created Table]0021010"
                // appears after last cell value record and before WINDOW2
                case 0x01BD:
                case 0x01C2:
                // Written by Excel 2007
                // rawData is multiple of 12 bytes long
                // appears after last cell value record and before WINDOW2 or drawing records
                //case 0x089D:
                case 0x089E:
                //case 0x08A7:

                //case 0x1001:
                //case 0x1006:
                //case 0x1007:
                //case 0x1009:
                //case 0x100A:
                //case 0x100B:
                case 0x100C:  //AttachedLabel
                //case 0x1014:  //ChartFormat
                //case 0x1017:  //Bar
                case 0x1018:  //Line
                //case 0x1019:  //Pie
                //case 0x101A:  //Area
                case 0x101B:  //Scatter
                //case 0x101D:  //Axis
                //case 0x101E:  //Tick
                //case 0x101F:  //ValueRange
                case 0x1020:// "CatSerRange";
                case 0x1021:// "AxisLine";
                case 0x1022:// "CrtLink";
                case 0x1024:// "DefaultText";
                //case 0x1025:  //Text
                case 0x1026:// "FontX";
                //case 0x1027:  //ObjectLink
                //case 0x1032:  //Frame
                case 0x1033:  //Begin
                case 0x1034:  //End
                case 0x1035:  //PlotArea
                //case 0x103A: //Chart3d
                case 0x1041: //AxisParent
                case 0x1043: //LegendException
                case 0x1044: //ShtProps
                case 0x1045:  //SerToCrt
                case 0x1046: //AxesUsed
                case 0x104A: //SerParent
                case 0x104B: //SerAuxTrend
                case 0x104E:  //IFmtRecord
                case 0x104F:  //Pos
                case 0x1051:  //BRAI
                case 0x105C:  //ClrtClient
                case 0x105D:  //SerFmt
                case 0x105F:  //Chart3DBarShape
                //case 0x1060:  //Fbi
                case 0x1062:  //AxcExt
                case 0x1063:  //Dat
                case 0x1064:  //PlotGrowth
                //case 0x1065:  //SIIndex
                case 0x1066:  //GelFrame
                case 4200:    //Fbi2
                case 4199:  //BopPopCustom
                case 4193://BopPop
                case 4176: //AlRuns
                    return true;
            }
            return false;
        }

        public override short Sid
        {
            get { return (short)_sid; }
        }


        /** Unlike the other Record.Clone methods this Is a shallow Clone*/
        public override Object Clone()
        {
            //UnknownRecord rec = new UnknownRecord();
            //rec._sid= _sid;
            //rec._rawData = _rawData;
            //return rec;

            // immutable - ok to return this
            return this;
        }
    }
}