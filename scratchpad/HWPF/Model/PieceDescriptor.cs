
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



namespace NPOI.HWPF.Model
{
    using System;
    using NPOI.Util;

    public class PieceDescriptor
    {

        short descriptor;
        private static BitField fNoParaLast = BitFieldFactory.GetInstance(0x01);
        private static BitField fPaphNil = BitFieldFactory.GetInstance(0x02);
        private static BitField fCopied = BitFieldFactory.GetInstance(0x04);
        int fc;
        PropertyModifier prm;
        bool unicode;


        public PieceDescriptor(byte[] buf, int offset)
        {
            descriptor = LittleEndian.GetShort(buf, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            fc = LittleEndian.GetInt(buf, offset);
            offset += LittleEndianConsts.INT_SIZE;
            prm = new PropertyModifier(LittleEndian.GetShort(buf, offset));

            // see if this piece uses unicode.
            if ((fc & 0x40000000) == 0)
            {
                unicode = true;
            }
            else
            {
                unicode = false;
                fc &= ~(0x40000000);//gives me FC in doc stream
                fc /= 2;
            }

        }

        public int FilePosition
        {
            get
            {
                return fc;
            }
            set 
            {
                fc = value;
            }
        }

        public bool IsUnicode
        {
            get
            {
                return unicode;
            }
        }

        internal byte[] ToByteArray()
        {
            // Set up the fc
            int tempFc = fc;
            if (!unicode)
            {
                tempFc *= 2;
                tempFc |= (0x40000000);
            }

            int offset = 0;
            byte[] buf = new byte[8];
            LittleEndian.PutShort(buf, offset, descriptor);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutInt(buf, offset, tempFc);
            offset += LittleEndianConsts.INT_SIZE;
            LittleEndian.PutShort(buf, offset, prm.GetValue());

            return buf;

        }

        public PropertyModifier Prm
        {
            get
            {
                return prm;
            }
        }


        public static int SizeInBytes
        {
            get
            {
                return 8;
            }
        }

        public override bool Equals(Object o)
        {
            PieceDescriptor pd = (PieceDescriptor)o;

            return descriptor == pd.descriptor && prm == pd.prm && unicode == pd.unicode;
        }
    }
}