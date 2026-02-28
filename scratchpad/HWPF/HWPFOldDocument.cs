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
using NPOI.Util;
using System.Text;
using NPOI.HWPF.Model;
using System;
using NPOI.HWPF.UserModel;
using NPOI.POIFS.FileSystem;
using System.IO;
namespace NPOI.HWPF
{

    /**
     * Provides very simple support for old (Word 6 / Word 95)
     *  files.
     */
    public class HWPFOldDocument : HWPFDocumentCore
    {
        private TextPieceTable tpt;

        public HWPFOldDocument(POIFSFileSystem fs)
            : this(fs.Root)
        {

        }
        [Obsolete]
        public HWPFOldDocument(DirectoryNode directory, POIFSFileSystem fs):this(directory)
        {
            
        }
        
        public HWPFOldDocument(DirectoryNode directory)
            : base(directory)
        {


            // Where are things?
            int sedTableOffset = LittleEndian.GetInt(_mainStream, 0x88);
            int sedTableSize = LittleEndian.GetInt(_mainStream, 0x8c);
            int chpTableOffset = LittleEndian.GetInt(_mainStream, 0xb8);
            int chpTableSize = LittleEndian.GetInt(_mainStream, 0xbc);
            int papTableOffset = LittleEndian.GetInt(_mainStream, 0xc0);
            int papTableSize = LittleEndian.GetInt(_mainStream, 0xc4);
            //int shfTableOffset = LittleEndian.GetInt(_mainStream, 0x60);
            //int shfTableSize   = LittleEndian.GetInt(_mainStream, 0x64);
            int complexTableOffset = LittleEndian.GetInt(_mainStream, 0x160);

            // We need to get hold of the text that Makes up the
            //  document, which might be regular or fast-saved
            StringBuilder text = new StringBuilder();
            if (_fib.IsFComplex())
            {
                ComplexFileTable cft = new ComplexFileTable(
                        _mainStream, _mainStream,
                        complexTableOffset, _fib.GetFcMin()
                );
                tpt = cft.GetTextPieceTable();

                foreach (TextPiece tp in tpt.TextPieces)
                {
                    text.Append(tp.GetStringBuilder());
                }
            }
            else
            {
                // TODO Discover if these older documents can ever hold Unicode Strings?
                //  (We think not, because they seem to lack a Piece table)
                // TODO Build the Piece Descriptor properly
                //  (We have to fake it, as they don't seem to have a proper Piece table)
                PieceDescriptor pd = new PieceDescriptor(new byte[] { 0, 0, 0, 0, 0, 127, 0, 0 }, 0);
                pd.FilePosition = _fib.GetFcMin();

                // Generate a single Text Piece Table, with a single Text Piece
                //  which covers all the (8 bit only) text in the file
                tpt = new TextPieceTable();
                byte[] textData = new byte[_fib.GetFcMac() - _fib.GetFcMin()];
                Array.Copy(_mainStream, _fib.GetFcMin(), textData, 0, textData.Length);
                TextPiece tp = new TextPiece(
                        0, textData.Length, textData, pd
                );
                tpt.Add(tp);

                text.Append(tp.GetStringBuilder());
            }

            _text = tpt.Text;

            // Now we can fetch the character and paragraph properties
            _cbt = new OldCHPBinTable(
                    _mainStream, chpTableOffset, chpTableSize,
                    _fib.GetFcMin(), tpt
            );
            _pbt = new OldPAPBinTable(
                    _mainStream, chpTableOffset, papTableSize,
                    _fib.GetFcMin(), tpt
            );
            _st = new OldSectionTable(
                    _mainStream, chpTableOffset, sedTableSize,
                    _fib.GetFcMin(), tpt
            );
        }
        public override Range GetOverallRange()
        {
            // Life is easy when we have no footers, headers or unicode!
            return new Range(0, _fib.GetFcMac() - _fib.GetFcMin(), this);
        }
        public override Range GetRange()
        {
            return GetOverallRange();
        }
        public override TextPieceTable TextTable
        {
            get
            {
                return tpt;
            }
        }
        private StringBuilder _text;
        public override StringBuilder Text
        {
            get
            {
                return _text;
            }
        }


        public override void Write(Stream out1)
        {
            throw new InvalidOperationException("Writing is not available for the older file formats");
        }
    }
}
