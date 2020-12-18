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

using System;
namespace NPOI.HWPF.Model
{
    /**
     * Normally PropertyNodes only ever work in characters, but
     *  a few cases actually store bytes, and this lets everything
     *  still work despite that.
     * It handles the conversion as required between bytes
     *  and characters.
     */
    public abstract class BytePropertyNode : PropertyNode
    {
        private int startBytes;
        private int endBytes;

        /**
         * @param fcStart The start of the text for this property, in _bytes_
         * @param fcEnd The end of the text for this property, in _bytes_
         */
        public BytePropertyNode(int fcStart, int fcEnd, CharIndexTranslator translator, object buf)
            : base(
                translator.GetCharIndex(fcStart),
                translator.GetCharIndex(fcEnd, translator.GetCharIndex(fcStart)), buf)
        {

            if (fcStart > fcEnd)
                throw new ArgumentException("fcStart (" + fcStart
                        + ") > fcEnd (" + fcEnd + ")");

            this.startBytes = fcStart;
            this.endBytes = fcEnd;           
        }
        public BytePropertyNode(int charStart, int charEnd, object buf)
            :base(charStart, charEnd, buf)
        {            

            if (charStart > charEnd)
                throw new ArgumentException("charStart (" + charStart
                        + ") > charEnd (" + charEnd + ")");

            this.startBytes = -1;
            this.endBytes = -1;
        }
        [Obsolete]
        public int StartBytes
        {
            get
            {
                return startBytes;
            }
        }
        [Obsolete]
        public int EndBytes
        {
            get
            {
                return endBytes;
            }
        }
    }
}