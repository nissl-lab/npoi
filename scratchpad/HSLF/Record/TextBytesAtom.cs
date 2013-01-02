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

namespace NPOI.HSLF.Record
{
    using System;
    using System.IO;
    using NPOI.Util;
    using System.Text;

    /**
     * A TextBytesAtom (type 4008). Holds text in ascii form (unknown
     *  code page, for now assumed to be the default of
     *  NPOI.UTIL.StringUtil, which is the Excel default).
     * The trailing return character is always stripped from this
     *
     * @author Nick Burch
     */

    public class TextBytesAtom : RecordAtom
    {
        private byte[] _header;
        private static long _type = 4008L;

        /** The bytes that make up the text */
        private byte[] _text;

        /** Grabs the text. Uses the default codepage */
        public String GetText()
        {
            return StringUtil.GetFromCompressedUnicode(_text, 0, _text.Length);
        }

        /** Updates the text in the Atom. Must be 8 bit ascii */
        public void SetText(byte[] b)
        {
            // Set the text
            _text = b;

            // Update the size (header bytes 5-8)
            LittleEndian.PutInt(_header, 4, _text.Length);
        }

        /* *************** record code follows ********************** */

        /**
         * For the TextBytes Atom
         */
        protected TextBytesAtom(byte[] source, int start, int len)
        {
            // Sanity Checking
            if (len < 8) { len = 8; }

            // Get the header
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Grab the text
            _text = new byte[len - 8];
            Array.Copy(source, start + 8, _text, 0, len - 8);
        }

        /**
         * Create an empty TextBytes Atom
         */
        public TextBytesAtom()
        {
            _header = new byte[8];
            LittleEndian.PutUShort(_header, 0, 0);
            LittleEndian.PutUShort(_header, 2, (int)_type);
            LittleEndian.PutInt(_header, 4, 0);

            _text = new byte[] { };
        }

        /**
         * We are of type 4008
         */
        public override long RecordType
        {
            get
            {
                return _type;
            }
        }

        /**
         * Write the contents of the record back, so it can be written
         *  to disk
         */
        public override void WriteOut(Stream out1)
        {
            // Header - size or type unChanged
            out1.Write(_header, (int)out1.Position, _header.Length);

            // Write out our text
            out1.Write(_text, (int)out1.Position, _text.Length);
        }

        /**
         * dump debug info; use GetText() to return a string
         * representation of the atom
         */
        public override String ToString()
        {
            StringBuilder out1 = new StringBuilder();
            out1.Append("TextBytesAtom:\n");
            out1.Append(HexDump.Dump(_text, 0, 0));
            return out1.ToString();
        }
    }
}

