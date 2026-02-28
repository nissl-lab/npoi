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
    using NPOI.HWPF.UserModel;
    using NPOI.HWPF.SPRM;

    /**
     * DANGER - works in bytes!
     *
     * Make sure you call GetStart() / GetEnd() when you want characters
     *  (normal use), but GetStartByte() / GetEndByte() when you're
     *  Reading in / writing out!
     *
     * @author Ryan Ackley
     */

    public class CHPX : BytePropertyNode
    {
        internal CHPX(int charStart, int charEnd, SprmBuffer buf):
            base(charStart, charEnd, buf)
        {
            
        }
        [Obsolete]
        public CHPX(int fcStart, int fcEnd, CharIndexTranslator translator, byte[] grpprl)
            : base(fcStart, translator.LookIndexBackward(fcEnd), translator, new SprmBuffer(grpprl))
        {

        }
        [Obsolete]
        public CHPX(int fcStart, int fcEnd, CharIndexTranslator translator, SprmBuffer buf)
            : base(fcStart, translator.LookIndexBackward(fcEnd), translator, buf)
        {

        }


        public byte[] GetGrpprl()
        {
            return ((SprmBuffer)_buf).ToByteArray();
        }

        public SprmBuffer GetSprmBuf()
        {
            return (SprmBuffer)_buf;
        }

        public CharacterProperties GetCharacterProperties(StyleSheet ss, short istd)
        {
            if (ss == null)
            {
                // TODO Fix up for Word 6/95
                return new CharacterProperties();
            }

            CharacterProperties baseStyle = ss.GetCharacterStyle(istd);
            if (baseStyle == null)
                baseStyle = new CharacterProperties();

            CharacterProperties props = 
                CharacterSprmUncompressor.UncompressCHP(baseStyle, GetGrpprl(), 0);
            return props;
        }

        public override String ToString()
        {
            return "CHPX from " + Start + " to " + End +
               " (in bytes " + StartBytes + " to " + EndBytes + ")";
        }
    }
}


