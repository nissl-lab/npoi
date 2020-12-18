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
    using NPOI.HWPF.Model.IO;
    using System;

    /**
     * FontTable or in MS terminology sttbfffn is a common data structure written in all
     * Word files. The sttbfffn is an sttbf where each string is an FFN structure instead
     * of pascal-style strings. An sttbf is a string Table stored in file. Thus sttbffn
     * is like an Sttbf with an array of FFN structures that stores the font name strings
     *
     * @author Praveen Mathew
     */
    public class FontTable
    {
        private short _stringCount;// how many strings are included in the string table
        private short _extraDataSz;// size in bytes of the extra data

        // Added extra facilitator members
        private int lcbSttbfffn;// count of bytes in sttbfffn
        private int fcSttbfffn;// table stream offset for sttbfffn

        // FFN structure Containing strings of font names
        private Ffn[] _fontNames = null;


        public FontTable(byte[] buf, int offset, int lcbSttbfffn)
        {
            this.lcbSttbfffn = lcbSttbfffn;
            this.fcSttbfffn = offset;

            _stringCount = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _extraDataSz = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;

            _fontNames = new Ffn[_stringCount]; //Ffn corresponds to a Pascal style String in STTBF.

            for (int i = 0; i < _stringCount; i++)
            {
                _fontNames[i] = new Ffn(buf, offset);
                offset += _fontNames[i].GetSize();
            }
        }

        public short GetStringCount()
        {
            return _stringCount;
        }

        public short GetExtraDataSz()
        {
            return _extraDataSz;
        }

        public Ffn[] GetFontNames()
        {
            return _fontNames;
        }

        public int GetSize()
        {
            return lcbSttbfffn;
        }

        public String GetMainFont(int chpFtc)
        {
            if (chpFtc >= _stringCount)
            {
                //Console.WriteLine("Mismatch in chpFtc with stringCount");
                return null;
            }

            return _fontNames[chpFtc].GetMainFontName();
        }

        public String GetAltFont(int chpFtc)
        {
            if (chpFtc >= _stringCount)
            {
                //Console.WriteLine("Mismatch in chpFtc with stringCount");
                return null;
            }

            return _fontNames[chpFtc].GetAltFontName();
        }

        public void SetStringCount(short stringCount)
        {
            this._stringCount = stringCount;
        }

        public void WriteTo(HWPFFileSystem sys)
        {
            HWPFStream tableStream = sys.GetStream("1Table");

            byte[] buf = new byte[LittleEndianConsts.SHORT_SIZE];
            LittleEndian.PutShort(buf, _stringCount);
            tableStream.Write(buf);
            LittleEndian.PutShort(buf, _extraDataSz);
            tableStream.Write(buf);

            for (int i = 0; i < _fontNames.Length; i++)
            {
                tableStream.Write(_fontNames[i].ToArray());
            }

        }

        public override bool Equals(Object o)
        {
            bool retVal = true;

            if (((FontTable)o).GetStringCount() == _stringCount)
            {
                if (((FontTable)o).GetExtraDataSz() == _extraDataSz)
                {
                    Ffn[] fontNamesNew = ((FontTable)o).GetFontNames();
                    for (int i = 0; i < _stringCount; i++)
                    {
                        if (!(_fontNames[i].Equals(fontNamesNew[i])))
                            retVal = false;
                    }
                }
                else
                    retVal = false;
            }
            else
                retVal = false;


            return retVal;
        }



    }
}




