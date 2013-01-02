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
     * A TextHeaderAtom  (type 3999). Holds information on what kind of
     *  text is Contained in the TextBytesAtom / TextCharsAtom that follows
     *  straight after
     *
     * @author Nick Burch
     */

    public class TextHeaderAtom : RecordAtom, ParentAwareRecord
    {
        private byte[] _header;
        private static long _type = 3999L;
        private RecordContainer parentRecord;

        public static int TITLE_TYPE = 0;
        public static int BODY_TYPE = 1;
        public static int NOTES_TYPE = 2;
        public static int OTHER_TYPE = 4;
        public static int CENTRE_BODY_TYPE = 5;
        public static int CENTER_TITLE_TYPE = 6;
        public static int HALF_BODY_TYPE = 7;
        public static int QUARTER_BODY_TYPE = 8;

        /** The kind of text it is */
        private int textType;

        public int GetTextType() { return textType; }
        public void SetTextType(int type) { textType = type; }

        public RecordContainer ParentRecord
        {
            get { return parentRecord; }
            set { this.parentRecord = value; }
        }
        
        

        /* *************** record code follows ********************** */

        /**
         * For the TextHeader Atom
         */
        protected TextHeaderAtom(byte[] source, int start, int len)
        {
            // Sanity Checking - we're always 12 bytes long
            if (len < 12)
            {
                len = 12;
                if (source.Length - start < 12)
                {
                    throw new Exception("Not enough data to form a TextHeaderAtom (always 12 bytes long) - found " + (source.Length - start));
                }
            }

            // Get the header
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Grab the type
            textType = LittleEndian.GetInt(source, start + 8);
        }

        /**
         * Create a new TextHeader Atom, for an unknown type of text
         */
        public TextHeaderAtom()
        {
            _header = new byte[8];
            LittleEndian.PutUShort(_header, 0, 0);
            LittleEndian.PutUShort(_header, 2, (int)_type);
            LittleEndian.PutInt(_header, 4, 4);

            textType = OTHER_TYPE;
        }

        /**
         * We are of type 3999
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

            // Write out our type
            WriteLittleEndian(textType, out1);
        }
    }


}