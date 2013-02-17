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
    using NPOI.HSSF.Util;

    using NPOI.SS.Util;

    /**
     * The <c>HyperlinkRecord</c> wraps an HLINK-record 
     *  from the Excel-97 format.
     * Supports only external links for now (eg http://) 
     *
     * @author      Mark Hissink Muller <a href="mailto:mark@hissinkmuller.nl">mark@hissinkmuller.nl</a>
     * @author      Yegor Kozlov (yegor at apache dot org)
     */
    public class HyperlinkRecord : StandardRecord
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(HyperlinkRecord));
        /**
         * Link flags
         */
        public const int HLINK_URL = 0x01;  // File link or URL.
        public const int HLINK_ABS = 0x02;  // Absolute path.
        public const int HLINK_LABEL = 0x14;  // Has label.
        public const int HLINK_PLACE = 0x08;  // Place in worksheet.
        private const int HLINK_TARGET_FRAME = 0x80;  // has 'target frame'
        private const int HLINK_UNC_PATH = 0x100;  // has UNC path

        public static readonly GUID STD_MONIKER = GUID.Parse("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B");
        public static readonly GUID URL_MONIKER = GUID.Parse("79EAC9E0-BAF9-11CE-8C82-00AA004BA90B");
        public static readonly GUID FILE_MONIKER = GUID.Parse("00000303-0000-0000-C000-000000000046");

        /**
         * Tail of a URL link
         */
        public static readonly byte[] URL_uninterpretedTail = HexRead.ReadFromString("79 58 81 F4  3B 1D 7F 48   AF 2C 82 5D  C4 85 27 63   00 00 00 00  A5 AB 00 00"); 
        /**
         * Tail of a file link
         */
        public static readonly byte[] FILE_uninterpretedTail = HexRead.ReadFromString("FF FF AD DE  00 00 00 00   00 00 00 00  00 00 00 00   00 00 00 00  00 00 00 00");
        private static readonly int TAIL_SIZE = FILE_uninterpretedTail.Length;
        
        public const short sid = 0x1b8;

        /** cell range of this hyperlink */
        private CellRangeAddress _range;

        /**
         * 16-byte GUID
         */
        private GUID _guid;


        /**
         * Some sort of options for file links.
         */
        private short _fileOpts;

        /**
         * Link options. Can include any of HLINK_* flags.
         */
        private int _linkOpts;

        /**
         * Test label
         */
        private String _label=string.Empty;
        private String _targetFrame=string.Empty;
        /**
         * Moniker. Makes sense only for URL and file links
         */
        private GUID _moniker;
        /** in 8:3 DOS format No Unicode string header,
         * always 8-bit characters, zero-terminated */
        private String _shortFilename=string.Empty;
        /** Link */
        private String _address=string.Empty;
        /**
         * Text describing a place in document.  In Excel UI, this is appended to the
         * address, (after a '#' delimiter).<br/>
         * This field is optional.  If present, the {@link #HLINK_PLACE} must be set.
         */
        private String _textMark=string.Empty;
        /**
         * Remaining bytes
         */
        private byte[] _uninterpretedTail;

        /**
         * Create a new hyperlink
         */
        public HyperlinkRecord()
        {

        }

        /**
         * Read hyperlink from input stream
         *
         * @param in the stream to Read from
         */
        public HyperlinkRecord(RecordInputStream in1)
        {
            _range = new CellRangeAddress(in1);

            // 16-byte GUID
            _guid = new GUID(in1);

            /*
             * streamVersion (4 bytes): An unsigned integer that specifies the version number
             * of the serialization implementation used to save this structure. This value MUST equal 2.
             */
            int streamVersion = in1.ReadInt();
            if (streamVersion != 0x00000002)
            {
                throw new RecordFormatException("Stream Version must be 0x2 but found " + streamVersion);
            }
            _linkOpts = in1.ReadInt();

            if ((_linkOpts & HLINK_LABEL) != 0)
            {
                int label_len = in1.ReadInt();
                _label = in1.ReadUnicodeLEString(label_len);
            }
            if ((_linkOpts & HLINK_TARGET_FRAME) != 0)
            {
                int len = in1.ReadInt();
                _targetFrame = in1.ReadUnicodeLEString(len);
            }
            if ((_linkOpts & HLINK_URL) != 0 && (_linkOpts & HLINK_UNC_PATH) != 0)
            {
                _moniker = null;
                int nChars = in1.ReadInt();
                _address = in1.ReadUnicodeLEString(nChars);
            }
            if ((_linkOpts & HLINK_URL) != 0 && (_linkOpts & HLINK_UNC_PATH) == 0)
            {
                _moniker = new GUID(in1);

                if (URL_MONIKER.Equals(_moniker))
                {
                    int length = in1.ReadInt();
                    /*
                     * The value of <code>length<code> be either the byte size of the url field
                     * (including the terminating NULL character) or the byte size of the url field plus 24.
                     * If the value of this field is set to the byte size of the url field,
                     * then the tail bytes fields are not present.
                     */
                    int remaining = in1.Remaining;
                    if (length == remaining)
                    {
                        int nChars = length / 2;
                        _address = in1.ReadUnicodeLEString(nChars);
                    }
                    else
                    {
                        int nChars = (length - TAIL_SIZE) / 2;
                        _address = in1.ReadUnicodeLEString(nChars);
                        /*
                         * TODO: make sense of the remaining bytes
                         * According to the spec they consist of:
                         * 1. 16-byte  GUID: This field MUST equal
                         *    {0xF4815879, 0x1D3B, 0x487F, 0xAF, 0x2C, 0x82, 0x5D, 0xC4, 0x85, 0x27, 0x63}
                         * 2. Serial version, this field MUST equal 0 if present.
                         * 3. URI Flags
                         */
                        _uninterpretedTail = ReadTail(URL_uninterpretedTail, in1);
                    }
                }
                else if (FILE_MONIKER.Equals(_moniker))
                {
                    _fileOpts = in1.ReadShort();

                    int len = in1.ReadInt();
                    _shortFilename = StringUtil.ReadCompressedUnicode(in1, len);
                    _uninterpretedTail = ReadTail(FILE_uninterpretedTail, in1);
                    int size = in1.ReadInt();
                    if (size > 0)
                    {
                        int charDataSize = in1.ReadInt();

                        //From the spec: An optional unsigned integer that MUST be 3 if present
                        // but some files has 4
                        int usKeyValue = in1.ReadUShort();
                        _address = StringUtil.ReadUnicodeLE(in1, charDataSize / 2);
                    }
                    else
                    {
                        _address = null;
                    }
                }
                else if (STD_MONIKER.Equals(_moniker))
                {
                    _fileOpts = in1.ReadShort();

                    int len = in1.ReadInt();

                    byte[] path_bytes = new byte[len];
                    in1.ReadFully(path_bytes);

                    _address = Encoding.UTF8.GetString(path_bytes);
                }
            }

            if ((_linkOpts & HLINK_PLACE) != 0)
            {
                int len = in1.ReadInt();
                _textMark = in1.ReadUnicodeLEString(len);
            }

            if (in1.Remaining > 0)
            {
                Console.WriteLine(HexDump.ToHex(in1.ReadRemainder()));
            }
        }
        private static byte[] ReadTail(byte[] expectedTail, ILittleEndianInput in1)
        {
            byte[] result = new byte[TAIL_SIZE];
            in1.ReadFully(result);
            //if (false)
            //{ // Quite a few examples in the unit tests which don't have the exact expected tail
            //    for (int i = 0; i < expectedTail.Length; i++)
            //    {
            //        if (expectedTail[i] != result[i])
            //        {
            //           logger.Log( POILogger.ERROR, "Mismatch in tail byte [" + i + "]"
            //                    + "expected " + (expectedTail[i] & 0xFF) + " but got " + (result[i] & 0xFF));
            //        }
            //    }
            //}
            return result;
        }
        private static void WriteTail(byte[] tail, ILittleEndianOutput out1)
        {
            out1.Write(tail);
        }

        /**
         * Return the column of the first cell that Contains the hyperlink
         *
         * @return the 0-based column of the first cell that Contains the hyperlink
         */
        public int FirstColumn
        {
            get{return _range.FirstColumn;}
            set{_range.FirstColumn = value;}
        }

        /**
         * Set the column of the last cell that Contains the hyperlink
         *
         * @return the 0-based column of the last cell that Contains the hyperlink
        */
        public int LastColumn
        {
            get{return _range.LastColumn;}
            set{_range.LastColumn= value;}
        }

        /**
         * Return the row of the first cell that Contains the hyperlink
         *
         * @return the 0-based row of the first cell that Contains the hyperlink
         */
        public int FirstRow
        {
           get{ return _range.FirstRow;}
            set{_range.FirstRow = value;}
        }

        /**
         * Return the row of the last cell that Contains the hyperlink
         *
         * @return the 0-based row of the last cell that Contains the hyperlink
         */
        public int LastRow
        {
            get { return _range.LastRow; }
            set { _range.LastRow = value; }
        }

        /**
         * Returns a 16-byte guid identifier. Seems to always equal {@link STD_MONIKER}
         *
         * @return 16-byte guid identifier
         */
        public GUID Guid
        {
            get
            {
                return _guid;
            }
        }

        /**
         * Returns a 16-byte moniker.
         *
         * @return 16-byte moniker
         */
        public GUID Moniker
        {
            get
            {
                return _moniker;
            }
        }
        private static String CleanString(String s)
        {
            if (s == null)
            {
                return null;
            }
            int idx = s.IndexOf('\u0000');
            if (idx < 0)
            {
                return s;
            }
            return s.Substring(0, idx);
        }
        private static String AppendNullTerm(String s)
        {
            if (s == null)
            {
                return null;
            }
            return s + '\u0000';
        }

        /**
         * Return text label for this hyperlink
         *
         * @return  text to Display
         */
        public String Label
        {
            get
            {
                return CleanString(_label);
            }
            set 
            {
                _label = AppendNullTerm(value);
            }
        }

        /**
         * Hypelink Address. Depending on the hyperlink type it can be URL, e-mail, patrh to a file, etc.
         *
         * @return  the Address of this hyperlink
         */
        public String Address
        {
            get
            {
                if ((_linkOpts & HLINK_URL) != 0 && _moniker!=null && FILE_MONIKER.Equals(_moniker))
                    return CleanString(_address != null ? _address : _shortFilename);
                else if ((_linkOpts & HLINK_PLACE) != 0)
                    return CleanString(_textMark);
                else
                    return CleanString(_address);
            }
            set
            {
                if ((_linkOpts & HLINK_URL) != 0 && _moniker != null && FILE_MONIKER.Equals(_moniker))
                    _shortFilename = AppendNullTerm(value);
                else if ((_linkOpts & HLINK_PLACE) != 0)
                    _textMark = AppendNullTerm(value);
                else
                    _address = AppendNullTerm(value);
            }
        }
        public String TextMark
        {
            get
            {
                return CleanString(_textMark);
            }
            set 
            {
                _textMark = AppendNullTerm(value);
            }
        }

        /**
         * Link options. Must be a combination of HLINK_* constants.
         */
        public int LinkOptions
        {
            get
            {
                return _linkOpts;
            }
        }
        public String TargetFrame
        {
            get
            {
                return CleanString(_targetFrame);
            }
        }
        public String ShortFilename
        {
            get
            {
                return CleanString(_shortFilename);
            }
            set 
            {
                _shortFilename = AppendNullTerm(value);
            }
        }
        /**
         * Label options
         */
        public int LabelOptions
        {
            get
            {
                return 2;
            }
        }

        /**
         * Options for a file link
         */
        public int FileOptions
        {
            get
            {
                return _fileOpts;
            }
        }


        public override short Sid
        {
            get { return HyperlinkRecord.sid; }
        }


        public override void Serialize(ILittleEndianOutput out1)
        {
            _range.Serialize(out1);

            _guid.Serialize(out1);
            out1.WriteInt(0x00000002); // TODO const
            out1.WriteInt(_linkOpts);

            if ((_linkOpts & HLINK_LABEL) != 0)
            {
                out1.WriteInt(_label.Length);
                StringUtil.PutUnicodeLE(_label, out1);
            }
            if ((_linkOpts & HLINK_TARGET_FRAME) != 0)
            {
                out1.WriteInt(_targetFrame.Length);
                StringUtil.PutUnicodeLE(_targetFrame, out1);
            }

            if ((_linkOpts & HLINK_URL) != 0 && (_linkOpts & HLINK_UNC_PATH) != 0)
            {
                out1.WriteInt(_address.Length);
                StringUtil.PutUnicodeLE(_address, out1);
            }
            if ((_linkOpts & HLINK_URL) != 0 && (_linkOpts & HLINK_UNC_PATH) == 0)
            {
                _moniker.Serialize(out1);
                if (_moniker != null && URL_MONIKER.Equals(_moniker))
                {
                    if (_uninterpretedTail == null) 
                    {
                        out1.WriteInt(_address.Length * 2);
                        StringUtil.PutUnicodeLE(_address, out1);
                    }
                    else
                    {
                        out1.WriteInt(_address.Length * 2 + TAIL_SIZE);
                        StringUtil.PutUnicodeLE(_address, out1);
                        WriteTail(_uninterpretedTail, out1);
                    }
                }
                else if (_moniker != null && FILE_MONIKER.Equals(_moniker))
                {
                    out1.WriteShort(_fileOpts);
                    out1.WriteInt(_shortFilename.Length);
                    StringUtil.PutCompressedUnicode(_shortFilename, out1);

                    WriteTail(_uninterpretedTail, out1);
                    if (string.IsNullOrEmpty(_address))
                    {
                        out1.WriteInt(0);
                    }
                    else
                    {
                        int addrLen = _address.Length * 2;
                        out1.WriteInt(addrLen + 6);
                        out1.WriteInt(addrLen);
                        out1.WriteShort(0x0003); // TODO const
                        StringUtil.PutUnicodeLE(_address, out1);
                    }
                }
            }
            if ((_linkOpts & HLINK_PLACE) != 0)
            {
                out1.WriteInt(_textMark.Length);
                StringUtil.PutUnicodeLE(_textMark, out1);

            }
        }

        protected override int DataSize
        {
            get
            {
                int size = 0;
                size += 2 + 2 + 2 + 2;  //rwFirst, rwLast, colFirst, colLast
                size += GUID.ENCODED_SIZE;
                size += 4;  //label_opts
                size += 4;  //_linkOpts
                if ((_linkOpts & HLINK_LABEL) != 0)
                {
                    size += 4;  //link Length
                    size += _label.Length * 2;
                }
                if ((_linkOpts & HLINK_TARGET_FRAME) != 0)
                {
                    size += 4;  // int nChars
                    size += _targetFrame.Length * 2;
                }
                if ((_linkOpts & HLINK_URL) != 0 && (_linkOpts & HLINK_UNC_PATH) != 0)
                {
                    size += 4;  // int nChars
                    size += _address.Length * 2;
                }
                if ((_linkOpts & HLINK_URL) != 0 && (_linkOpts & HLINK_UNC_PATH) == 0)
                {
                    size += GUID.ENCODED_SIZE;  //moniker Length
                    if (_moniker!=null&&URL_MONIKER.Equals(_moniker))
                    {
                        size += 4;  //Address Length
                        size += _address.Length * 2;
                        if (_uninterpretedTail != null)
                        {
                            size += TAIL_SIZE;
                        }
                    }
                    else if (_moniker != null && FILE_MONIKER.Equals(_moniker))
                    {
                        size += 2;  //_fileOpts
                        size += 4;  //Address Length
                        size += _shortFilename == null ? 0 : _shortFilename.Length;
                        size += TAIL_SIZE;
                        size += 4;
                        if (!string.IsNullOrEmpty(_address))
                        {
                            size += 6;
                            size += _address.Length * 2;
                        }
                    }
                }
                if ((_linkOpts & HLINK_PLACE) != 0)
                {
                    size += 4;  //Address Length
                    size += _textMark.Length * 2;
                }
                return size;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[HYPERLINK RECORD]\n");
            buffer.Append("    .range            = ").Append(_range.FormatAsString()).Append("\n");
            buffer.Append("    .guid        = ").Append(_guid.FormatAsString()).Append("\n");
            buffer.Append("    .linkOpts          = ").Append(HexDump.IntToHex(this._linkOpts)).Append("\n");
            buffer.Append("    .label          = ").Append(Label).Append("\n");
            if ((_linkOpts & HLINK_TARGET_FRAME) != 0)
            {
                buffer.Append("    .targetFrame= ").Append(TargetFrame).Append("\n");
            }
            if((_linkOpts & HLINK_URL) != 0 && _moniker != null) 
            {
                buffer.Append("    .moniker          = ").Append(_moniker.FormatAsString()).Append("\n");
            }
            if ((_linkOpts & HLINK_PLACE) != 0) 
            {
                buffer.Append("    .targetFrame= ").Append(TextMark).Append("\n");
            }
            buffer.Append("    .address            = ").Append(Address).Append("\n");
            buffer.Append("[/HYPERLINK RECORD]\n");
            return buffer.ToString();
        }

        /// <summary>
        /// Initialize a new url link
        /// </summary>        
        public void CreateUrlLink()
        {
            _range = new CellRangeAddress(0, 0, 0, 0); 
            _guid = STD_MONIKER;
            
            _linkOpts = HLINK_URL | HLINK_ABS | HLINK_LABEL;
            Label = "";
            _moniker = URL_MONIKER;
            Address = "";
            _uninterpretedTail = URL_uninterpretedTail;
        }

        /// <summary>
        /// Initialize a new file link
        /// </summary>
        public void CreateFileLink()
        {
            _range = new CellRangeAddress(0, 0, 0, 0);
            _guid = STD_MONIKER;
            _linkOpts = HLINK_URL | HLINK_LABEL;
            _fileOpts = 0;
            Label = "";
            _moniker = FILE_MONIKER;
            Address= null;
            ShortFilename = "";
            _uninterpretedTail = FILE_uninterpretedTail;
        }

        /// <summary>
        /// Initialize a new document link
        /// </summary>
        public void CreateDocumentLink()
        {
            _range = new CellRangeAddress(0, 0, 0, 0);
            _guid = STD_MONIKER;
            _linkOpts = HLINK_LABEL | HLINK_PLACE;
            Label = "";
            _moniker = FILE_MONIKER;
            Address = "";
            TextMark = "";
        }

        public override Object Clone()
        {
            HyperlinkRecord rec = new HyperlinkRecord();
            rec._range = _range.Copy();
            rec._guid = _guid;
            
            rec._linkOpts = _linkOpts;
            rec._fileOpts = _fileOpts;
            rec._label = _label;
            rec._address = _address;
            rec._moniker = _moniker;
            rec._shortFilename = _shortFilename;
            rec._targetFrame = _targetFrame;
            rec._textMark = _textMark;
            rec._uninterpretedTail = _uninterpretedTail;
            return rec;
        }


    }
}