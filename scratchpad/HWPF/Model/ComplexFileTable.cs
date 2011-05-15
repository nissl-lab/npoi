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


namespace NPOI.HWPF.Model
{
    using NPOI.Util;
    using System.IO;
    using NPOI.HWPF.Model.IO;

    public class ComplexFileTable
    {

        private static byte GRPPRL_TYPE = 1;
        private static byte TEXT_PIECE_TABLE_TYPE = 2;

        protected TextPieceTable _tpt;

        public ComplexFileTable()
        {
            _tpt = new TextPieceTable();
        }

        public ComplexFileTable(byte[] documentStream, byte[] tableStream, int offset, int fcMin)
        {
            //skips through the prms before we reach the piece table. These contain data
            //for actual fast saved files
            while (tableStream[offset] == GRPPRL_TYPE)
            {
                offset++;
                int skip = LittleEndian.GetShort(tableStream, offset);
                offset += LittleEndianConstants.SHORT_SIZE + skip;
            }
            if (tableStream[offset] != TEXT_PIECE_TABLE_TYPE)
            {
                throw new IOException("The text piece table is corrupted");
            }
            int pieceTableSize = LittleEndian.GetInt(tableStream, ++offset);
            offset += LittleEndianConstants.INT_SIZE;
            _tpt = new TextPieceTable(documentStream, tableStream, offset, pieceTableSize, fcMin);
        }

        public TextPieceTable GetTextPieceTable()
        {
            return _tpt;
        }

        public void WriteTo(HWPFFileSystem sys)
        {
            HWPFStream docStream = sys.GetStream("WordDocument");
            HWPFStream tableStream = sys.GetStream("1Table");

            tableStream.Write(TEXT_PIECE_TABLE_TYPE);

            byte[] table = _tpt.WriteTo(docStream);

            byte[] numHolder = new byte[LittleEndianConstants.INT_SIZE];
            LittleEndian.PutInt(numHolder, table.Length);
            tableStream.Write(numHolder);
            tableStream.Write(table);
        }

    }
}

