/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.UserModel
{
    using System;
    using NPOI.Util;
    using NPOI.HWPF.Model;

    /**
     * This class Is used to determine line spacing for a paragraph.
     *
     * @author Ryan Ackley
     */
    public class LineSpacingDescriptor:BaseObject
    {
        short _dyaLine;
        short _fMultiLinespace;

        public LineSpacingDescriptor()
        {
        }

        public LineSpacingDescriptor(byte[] buf, int offset)
        {
            _dyaLine = LittleEndian.GetShort(buf, offset);
            _fMultiLinespace = LittleEndian.GetShort(buf, offset + LittleEndianConsts.SHORT_SIZE);
        }

        public void SetMultiLinespace(short fMultiLinespace)
        {
            _fMultiLinespace = fMultiLinespace;
        }

        public int ToInt()
        {
            byte[] intHolder = new byte[4];
            Serialize(intHolder, 0);
            return LittleEndian.GetInt(intHolder);
        }

        public void Serialize(byte[] buf, int offset)
        {
            LittleEndian.PutShort(buf, offset, _dyaLine);
            LittleEndian.PutShort(buf, offset + LittleEndianConsts.SHORT_SIZE, _fMultiLinespace);
        }

        public void SetDyaLine(short dyaLine)
        {
            _dyaLine = dyaLine;
        }
        public override bool Equals(Object o)
        {
            LineSpacingDescriptor lspd = (LineSpacingDescriptor)o;

            return _dyaLine == lspd._dyaLine && _fMultiLinespace == lspd._fMultiLinespace;
        }
    }
}