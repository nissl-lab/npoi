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

namespace NPOI.HSSF.Record.Crypto
{
    using System;
    using System.Text;
    using NPOI.Util;

    /**
     * Simple implementation of the alleged RC4 algorithm.
     *
     * Inspired by <A HREF="http://en.wikipedia.org/wiki/RC4">wikipedia's RC4 article</A>
     *
     * @author Josh Micich
     */
    internal class RC4
    {

        private int _i, _j;
        private byte[] _s = new byte[256];

        public RC4(byte[] key)
        {
            int key_length = key.Length;

            for (int i = 0; i < 256; i++)
                _s[i] = (byte)i;

            for (int i = 0, j = 0; i < 256; i++)
            {
                byte temp;

                j = (j + key[i % key_length] + _s[i]) & 255;
                temp = _s[i];
                _s[i] = _s[j];
                _s[j] = temp;
            }

            _i = 0;
            _j = 0;
        }

        public byte Output()
        {
            byte temp;
            _i = (_i + 1) & 255;
            _j = (_j + _s[_i]) & 255;

            temp = _s[_i];
            _s[_i] = _s[_j];
            _s[_j] = temp;

            return _s[(_s[_i] + _s[_j]) & 255];
        }

        public void Encrypt(byte[] in1)
        {
            for (int i = 0; i < in1.Length; i++)
            {
                in1[i] = (byte)(in1[i] ^ Output());
            }
        }
        public void Encrypt(byte[] in1, int OffSet, int len)
        {
            int end = OffSet + len;
            for (int i = OffSet; i < end; i++)
            {
                in1[i] = (byte)(in1[i] ^ Output());
            }

        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(GetType().Name).Append(" [");
            sb.Append("i=").Append(_i);
            sb.Append(" j=").Append(_j);
            sb.Append("]");
            sb.Append("\n");
            sb.Append(HexDump.Dump(_s, 0, 0));

            return sb.ToString();
        }
    }


}