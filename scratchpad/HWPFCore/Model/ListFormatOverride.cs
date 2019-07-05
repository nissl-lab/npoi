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

    public class ListFormatOverride
    {
        int _lsid;
        int _reserved1;
        int _reserved2;
        byte _clfolvl;
        byte[] _reserved3 = new byte[3];
        ListFormatOverrideLevel[] _levelOverrides;

        public ListFormatOverride(int lsid)
        {
            _lsid = lsid;
            _levelOverrides = new ListFormatOverrideLevel[0];
        }

        public ListFormatOverride(byte[] buf, int offset)
        {
            _lsid = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _reserved1 = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _reserved2 = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            _clfolvl = buf[offset++];
            Array.Copy(buf, offset, _reserved3, 0, _reserved3.Length);
            _levelOverrides = new ListFormatOverrideLevel[_clfolvl];
        }

        public int numOverrides()
        {
            return _clfolvl;
        }

        public int GetLsid()
        {
            return _lsid;
        }

        internal void SetLsid(int lsid)
        {
            _lsid = lsid;
        }

        public ListFormatOverrideLevel[] GetLevelOverrides()
        {
            return _levelOverrides;
        }

        public void SetOverride(int index, ListFormatOverrideLevel lfolvl)
        {
            _levelOverrides[index] = lfolvl;
        }

        public ListFormatOverrideLevel GetOverrideLevel(int level)
        {

            ListFormatOverrideLevel retLevel = null;

            for (int x = 0; x < _levelOverrides.Length; x++)
            {
                if (_levelOverrides[x].GetLevelNum() == level)
                {
                    retLevel = _levelOverrides[x];
                }
            }
            return retLevel;
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
            {
                return false;
            }

            ListFormatOverride lfo = (ListFormatOverride)obj;
            return lfo._clfolvl == _clfolvl && lfo._lsid == _lsid &&
              lfo._reserved1 == _reserved1 && lfo._reserved2 == _reserved2 &&
              Arrays.Equals(lfo._reserved3, _reserved3) &&
              Arrays.Equals(lfo._levelOverrides, _levelOverrides);
        }

        public byte[] ToArray()
        {
            byte[] buf = new byte[16];
            int offset = 0;
            LittleEndian.PutInt(buf, offset, _lsid);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, _reserved1);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(buf, offset, _reserved2);
            offset += LittleEndianConsts.INT_SIZE;
            buf[offset++] = _clfolvl;
            Array.Copy(_reserved3, 0, buf, offset, 3);

            return buf;
        }
    }
}

