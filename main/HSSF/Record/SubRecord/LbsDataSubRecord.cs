using System;
using System.Text;
using NPOI.Util;
using NPOI.SS.Formula.PTG;
using System.Globalization;

namespace NPOI.HSSF.Record
{
    public class LbsDataSubRecord : SubRecord
    {

        public const int sid = 0x0013;
        /**
          * From [MS-XLS].pdf 2.5.147 FtLbsData:
          *
          * An unsigned integer that indirectly specifies whether
          * some of the data in this structure appear in a subsequent Continue record.
          * If _cbFContinued is 0x00, all of the fields in this structure except sid and _cbFContinued
          *  MUST NOT exist. If this entire structure is Contained within the same record,
          * then _cbFContinued MUST be greater than or equal to the size, in bytes,
          * of this structure, not including the four bytes for the ft and _cbFContinued fields
          */
        private int _cbFContinued;

        /**
         * a formula that specifies the range of cell values that are the items in this list.
         */
        private int _unknownPreFormulaInt;
        private Ptg _linkPtg;
        private Byte? _unknownPostFormulaByte;

        /**
         * An unsigned integer that specifies the number of items in the list.
         */
        private int _cLines;

        /**
         * An unsigned integer that specifies the one-based index of the first selected item in this list.
         * A value of 0x00 specifies there is no currently selected item.
         */
        private int _iSel;

        /**
         *  flags that tell what data follows
         */
        private int _flags;

        /**
         * An ObjId that specifies the edit box associated with this list.
         * A value of 0x00 specifies that there is no edit box associated with this list.
         */
        private int _idEdit;

        /**
         * An optional LbsDropData that specifies properties for this dropdown control.
         * This field MUST exist if and only if the Containing Obj?s cmo.ot is equal to 0x14.
         */
        private LbsDropData _dropData;

        /**
         * An optional array of strings where each string specifies an item in the list.
         * The number of elements in this array, if it exists, MUST be {@link #_cLines}
         */
        private String[] _rgLines;

        /**
         * An optional array of bools that specifies
         * which items in the list are part of a multiple selection
         */
        private bool[] _bsels;

        LbsDataSubRecord()
        {

        }
    /**
     * @param in the stream to read data from
     * @param cbFContinued the seconf short in the record header
     * @param cmoOt the Containing Obj's {@link CommonObjectDataSubRecord#field_1_objectType}
     */
        public LbsDataSubRecord(ILittleEndianInput in1, int cbFContinued, int cmoOt)
        {
            _cbFContinued = cbFContinued;

            int encodedTokenLen = in1.ReadUShort();
            if (encodedTokenLen > 0)
            {
                int formulaSize = in1.ReadUShort();
                _unknownPreFormulaInt = in1.ReadInt();

                Ptg[] ptgs = Ptg.ReadTokens(formulaSize, in1);
                if (ptgs.Length != 1)
                {
                    throw new RecordFormatException("Read " + ptgs.Length
                            + " tokens but expected exactly 1");
                }
                _linkPtg = ptgs[0];
                switch (encodedTokenLen - formulaSize - 6)
                {
                    case 1:
                        _unknownPostFormulaByte = (byte)in1.ReadByte();
                        break;
                    case 0:
                        _unknownPostFormulaByte = null;
                        break;
                    default:
                        throw new RecordFormatException("Unexpected leftover bytes");
                }
            }

            _cLines = in1.ReadUShort();
            _iSel = in1.ReadUShort();
            _flags = in1.ReadUShort();
            _idEdit = in1.ReadUShort();

            // From [MS-XLS].pdf 2.5.147 FtLbsData:
            // This field MUST exist if and only if the Containing Obj?s cmo.ot is equal to 0x14.
            if (cmoOt == 0x14)
            {
                _dropData = new LbsDropData(in1);
            }

            // From [MS-XLS].pdf 2.5.147 FtLbsData:
            // This array MUST exist if and only if the fValidPlex flag (0x2) is set
            if ((_flags & 0x2) != 0)
            {
                _rgLines = new String[_cLines];
                for (int i = 0; i < _cLines; i++)
                {
                    _rgLines[i] = StringUtil.ReadUnicodeString(in1);
                }
            }

            // bits 5-6 in the _flags specify the type
            // of selection behavior this list control is expected to support

            // From [MS-XLS].pdf 2.5.147 FtLbsData:
            // This array MUST exist if and only if the wListType field is not equal to 0.
            if (((_flags >> 4) & 0x2) != 0)
            {
                _bsels = new bool[_cLines];
                for (int i = 0; i < _cLines; i++)
                {
                    _bsels[i] = in1.ReadByte() == 1;
                }
            }

        }
        public override bool IsTerminating
        {
            get
            {
                return true;
            }
        }
        /**
 *
 * @return the formula that specifies the range of cell values that are the items in this list.
 */
        public Ptg Formula
        {
            get
            {
                return _linkPtg;
            }
        }

        /**
         * @return the number of items in the list
         */
        public int NumberOfItems
        {
            get
            {
                return _cLines;
            }
        }
        public override short Sid
        {
            get { return sid; }
        }
        public override int DataSize
        {
            get
            {
                int result = 2; // 2 Initial shorts

                // optional link formula
                if (_linkPtg != null)
                {
                    result += 2; // encoded Ptg size
                    result += 4; // unknown int
                    result += _linkPtg.Size;
                    if (_unknownPostFormulaByte != null)
                    {
                        result += 1;
                    }
                }

                result += 4 * 2; // 4 shorts
                if (_dropData != null)
                {
                    result += _dropData.DataSize;
                }
                if (_rgLines != null)
                {
                    foreach (String str in _rgLines)
                    {
                        result += StringUtil.GetEncodedSize(str);
                    }
                }
                if (_bsels != null)
                {
                    result += _bsels.Length;
                }
                return result;
            }
        }
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(_cbFContinued); // note - this is *not* the size
            if (_linkPtg == null)
            {
                out1.WriteShort(0);
            }
            else
            {
                int formulaSize = _linkPtg.Size;
                int linkSize = formulaSize + 6;
                if (_unknownPostFormulaByte != null)
                {
                    linkSize++;
                }
                out1.WriteShort(linkSize);
                out1.WriteShort(formulaSize);
                out1.WriteInt(_unknownPreFormulaInt);
                _linkPtg.Write(out1);
                if (_unknownPostFormulaByte != null)
                {
                    out1.WriteByte(Convert.ToByte(_unknownPostFormulaByte, CultureInfo.InvariantCulture));
                }
            }
            out1.WriteShort(_cLines);
            out1.WriteShort(_iSel);
            out1.WriteShort(_flags);
            out1.WriteShort(_idEdit);

            if (_dropData != null)
            {
                _dropData.Serialize(out1);
            }

            if (_rgLines != null)
            {
                foreach (String str in _rgLines)
                {
                    StringUtil.WriteUnicodeString(out1, str);
                }
            }

            if (_bsels != null)
            {
                foreach (bool val in _bsels)
                {
                    out1.WriteByte(val ? 1 : 0);
                }
            }
        }
        private static Ptg ReadRefPtg(byte[] formulaRawBytes)
        {
            ILittleEndianInput in1 = new LittleEndianByteArrayInputStream(formulaRawBytes);
            byte ptgSid = (byte)in1.ReadByte();
            switch (ptgSid)
            {
                case AreaPtg.sid: return new AreaPtg(in1);
                case Area3DPtg.sid: return new Area3DPtg(in1);
                case RefPtg.sid: return new RefPtg(in1);
                case Ref3DPtg.sid: return new Ref3DPtg(in1);
            }
            return null;
        }
        public override Object Clone()
        {
            return this;
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(256);

            sb.Append("[ftLbsData]\n");
            sb.Append("    .unknownshort1 =").Append(HexDump.ShortToHex(_cbFContinued)).Append("\n");
            sb.Append("    .formula        = ").Append('\n');
            sb.Append(_linkPtg.ToString()).Append(_linkPtg.RVAType).Append('\n');
            sb.Append("    .nEntryCount   =").Append(HexDump.ShortToHex(_cLines)).Append("\n");
            sb.Append("    .selEntryIx    =").Append(HexDump.ShortToHex(_iSel)).Append("\n");
            sb.Append("    .style         =").Append(HexDump.ShortToHex(_flags)).Append("\n");
            sb.Append("    .unknownshort10=").Append(HexDump.ShortToHex(_idEdit)).Append("\n");
            if (_dropData != null) sb.Append('\n').Append(_dropData.ToString());
            sb.Append("[/ftLbsData]\n");
            return sb.ToString();
        }

        /**
  *
  * @return a new instance of LbsDataSubRecord to construct auto-filters
  * @see org.apache.poi.hssf.model.ComboboxShape#createObjRecord(org.apache.poi.hssf.usermodel.HSSFSimpleShape, int)
  */
        public static LbsDataSubRecord CreateAutoFilterInstance()
        {
            LbsDataSubRecord lbs = new LbsDataSubRecord();
            lbs._cbFContinued = 0x1FEE;  //autofilters seem to alway have this magic number
            lbs._iSel = 0x000;

            lbs._flags = 0x0301;
            lbs._dropData = new LbsDropData();
            lbs._dropData._wStyle = LbsDropData.STYLE_COMBO_SIMPLE_DROPDOWN;

            // the number of lines to be displayed in the dropdown
            lbs._dropData._cLine = 8;
            return lbs;
        }
    }
 
        /**
     * This structure specifies properties of the dropdown list control
     */
    public class LbsDropData {

        /**
 * Combo dropdown control
 */
        public const int STYLE_COMBO_DROPDOWN = 0;
        /**
         * Combo Edit dropdown control
         */
        public const int STYLE_COMBO_EDIT_DROPDOWN = 1;
        /**
         * Simple dropdown control (just the dropdown button)
         */
        public const int STYLE_COMBO_SIMPLE_DROPDOWN = 2;
        /**
         *  An unsigned integer that specifies the style of this dropdown. 
         */
        internal int _wStyle;

        /**
         * An unsigned integer that specifies the number of lines to be displayed in the dropdown.
         */
        internal int _cLine;

        /**
         * An unsigned integer that specifies the smallest width in pixels allowed for the dropdown window
         */
        private int _dxMin;

        /**
         * a string that specifies the current string value in the dropdown
         */
        private String _str;

        /**
         * Optional, undefined and MUST be ignored.
         * This field MUST exist if and only if the size of str in bytes is an odd number
         */
        private Byte _unused;

        public LbsDropData()
        {
            _str = "";
            _unused = 0;
        }

        public LbsDropData(ILittleEndianInput in1){
            _wStyle = in1.ReadUShort();
            _cLine = in1.ReadUShort();
            _dxMin = in1.ReadUShort();
            _str = StringUtil.ReadUnicodeString(in1);
            if(StringUtil.GetEncodedSize(_str) % 2 != 0){
                _unused = (byte)in1.ReadByte();
            }
        }

        public void Serialize(ILittleEndianOutput out1) {
            out1.WriteShort(_wStyle);
            out1.WriteShort(_cLine);
            out1.WriteShort(_dxMin);
            StringUtil.WriteUnicodeString(out1, _str);
            out1.WriteByte(_unused);
        }

        public int DataSize
        {
            get
            {
                int size = 6;
                size += StringUtil.GetEncodedSize(_str);
                size += _unused;
                return size;
            }
        }

        public override String ToString(){
            StringBuilder sb = new StringBuilder();
            sb.Append("[LbsDropData]\n");
            sb.Append("  ._wStyle:  ").Append(_wStyle).Append('\n');
            sb.Append("  ._cLine:  ").Append(_cLine).Append('\n');
            sb.Append("  ._dxMin:  ").Append(_dxMin).Append('\n');
            sb.Append("  ._str:  ").Append(_str).Append('\n');
            sb.Append("  ._unused:  ").Append(_unused).Append('\n');
            sb.Append("[/LbsDropData]\n");

            return sb.ToString();
        }
    }
}
