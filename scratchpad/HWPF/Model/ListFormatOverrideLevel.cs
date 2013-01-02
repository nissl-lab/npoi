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
    using System;
    using NPOI.Util;

    public class ListFormatOverrideLevel
    {
        private static int BASE_SIZE = 8;

        int _iStartAt;
        byte _info;
        private static BitField _ilvl = BitFieldFactory.GetInstance(0xf);
        private static BitField _fStartAt = BitFieldFactory.GetInstance(0x10);
        private static BitField _fFormatting = BitFieldFactory.GetInstance(0x20);
        byte[] _reserved = new byte[3];
        ListLevel _lvl;

        public ListFormatOverrideLevel(byte[] buf, int offset)
        {
            _iStartAt = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _info = buf[offset++];
            Array.Copy(buf, offset, _reserved, 0, _reserved.Length);
            offset += _reserved.Length;

            if (_fFormatting.GetValue(_info) > 0)
            {
                _lvl = new ListLevel(buf, offset);
            }
        }

        public ListLevel GetLevel()
        {
            return _lvl;
        }

        public int GetLevelNum()
        {
            return _ilvl.GetValue(_info);
        }

        public bool IsFormatting()
        {
            return _fFormatting.GetValue(_info) != 0;
        }

        public bool IsStartAt()
        {
            return _fStartAt.GetValue(_info) != 0;
        }

        public int GetSizeInBytes()
        {
            return (_lvl == null ? BASE_SIZE : BASE_SIZE + _lvl.GetSizeInBytes());
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
            {
                return false;
            }
            ListFormatOverrideLevel lfolvl = (ListFormatOverrideLevel)obj;
            bool lvlEquality = false;
            if (_lvl != null)
            {
                lvlEquality = _lvl.Equals(lfolvl._lvl);
            }
            else
            {
                lvlEquality = lfolvl._lvl == null;
            }

            return lvlEquality && lfolvl._iStartAt == _iStartAt && lfolvl._info == _info &&
              Arrays.Equals(lfolvl._reserved, _reserved);
        }

        public byte[] ToArray()
        {
            byte[] buf = new byte[GetSizeInBytes()];

            int offset = 0;
            LittleEndian.PutInt(buf, _iStartAt);
            offset += LittleEndianConsts.INT_SIZE;
            buf[offset++] = _info;
            Array.Copy(_reserved, 0, buf, offset, 3);
            offset += 3;

            if (_lvl != null)
            {
                byte[] levelBuf = _lvl.ToArray();
                Array.Copy(levelBuf, 0, buf, offset, levelBuf.Length);
            }

            return buf;
        }

        public int GetIStartAt()
        {
            return _iStartAt;
        }
    }
}


