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
    using System;

    public class ListData
    {
        private int _lsid;
        private int _tplc;
        private short[] _rgistd;
        private byte _info;
        private static BitField _fSimpleList = BitFieldFactory.GetInstance(0x1);
        private static BitField _fRestartHdn = BitFieldFactory.GetInstance(0x2);
        private byte _reserved;
        ListLevel[] _levels;

        public ListData(int listID, bool numbered)
        {
            _lsid = listID;
            _rgistd = new short[9];

            for (int x = 0; x < 9; x++)
            {
                _rgistd[x] = (short)StyleSheet.NIL_STYLE;
            }

            _levels = new ListLevel[9];

            for (int x = 0; x < _levels.Length; x++)
            {
                _levels[x] = new ListLevel(x, numbered);
            }
        }

        public ListData(byte[] buf, int offset)
        {
            _lsid = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _tplc = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _rgistd = new short[9];
            for (int x = 0; x < 9; x++)
            {
                _rgistd[x] = LittleEndian.GetShort(buf, offset);
                offset += LittleEndianConsts.SHORT_SIZE;
            }
            _info = buf[offset++];
            _reserved = buf[offset];
            if (_fSimpleList.GetValue(_info) > 0)
            {
                _levels = new ListLevel[1];
            }
            else
            {
                _levels = new ListLevel[9];
            }

        }

        public int GetLsid()
        {
            return _lsid;
        }

        public int numLevels()
        {
            return _levels.Length;
        }

        public void SetLevel(int index, ListLevel level)
        {
            _levels[index] = level;
        }

        public ListLevel[] GetLevels()
        {
            return _levels;
        }

        /**
         * Gets the level associated to a particular List at a particular index.
         *
         * @param index 1-based index
         * @return a list level
         */
        public ListLevel GetLevel(int index)
        {
            return _levels[index - 1];
        }

        public int GetLevelStyle(int index)
        {
            return _rgistd[index];
        }

        public void SetLevelStyle(int index, int styleIndex)
        {
            _rgistd[index] = (short)styleIndex;
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
            {
                return false;
            }

            ListData lst = (ListData)obj;
            return lst._info == _info && Arrays.Equals(lst._levels, _levels) &&
              lst._lsid == _lsid && lst._reserved == _reserved && lst._tplc == _tplc &&
              Arrays.Equals(lst._rgistd, _rgistd);
        }

        internal int ResetListID()
        {
            _lsid = (int)((new Random((int)DateTime.Now.Ticks)).Next(0,100)/100 * DateTime.Now.Millisecond);
            return _lsid;
        }

        public byte[] ToArray()
        {
            byte[] buf = new byte[28];
            int offset = 0;
            LittleEndian.PutInt(buf, _lsid);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, _tplc);
            offset += LittleEndianConsts.INT_SIZE;
            for (int x = 0; x < 9; x++)
            {
                LittleEndian.PutShort(buf, offset, _rgistd[x]);
                offset += LittleEndianConsts.SHORT_SIZE;
            }
            buf[offset++] = _info;
            buf[offset] = _reserved;
            return buf;
        }
    }
}


