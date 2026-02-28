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
     * This class Is used to represent a date and time in a Word document.
     *
     * @author Ryan Ackley
     */
    public class DateAndTime : BaseObject
    {
        public static int SIZE = 4;
        private short _info;
        private static BitField _minutes = BitFieldFactory.GetInstance(0x3f);
        private static BitField _hours = BitFieldFactory.GetInstance(0x7c0);
        private static BitField _dom = BitFieldFactory.GetInstance(0xf800);
        private short _info2;
        private static BitField _months = BitFieldFactory.GetInstance(0xf);
        private static BitField _years = BitFieldFactory.GetInstance(0x1ff0);
        private static BitField _weekday = BitFieldFactory.GetInstance(0xe000);

        public DateAndTime()
        {
        }

        public DateAndTime(byte[] buf, int offset)
        {
            _info = LittleEndian.GetShort(buf, offset);
            _info2 = LittleEndian.GetShort(buf, offset + LittleEndianConsts.SHORT_SIZE);
        }

        public void Serialize(byte[] buf, int offset)
        {
            LittleEndian.PutShort(buf, offset, _info);
            LittleEndian.PutShort(buf, offset + LittleEndianConsts.SHORT_SIZE, _info2);
        }

        public override bool Equals(Object o)
        {
            DateAndTime dttm = (DateAndTime)o;
            return _info == dttm._info && _info2 == dttm._info2;
        }

    }
}