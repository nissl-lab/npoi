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

    /**
     * This class represents a comment on a slide, in the format used by
     *  PPT 2000/XP/etc. (PPT 97 uses plain Escher Text Boxes for comments)
     * @author Nick Burch
     */
    public class Comment2000 : RecordContainer
    {
        private byte[] _header;
        private static long _type = 12000;

        // Links to our more interesting children

        /**
         * An optional string that specifies the name of the author of the presentation comment.
         */
        private CString authorRecord;
        /**
         * An optional string record that specifies the text of the presentation comment
         */
        private CString authorInitialsRecord;
        /**
         * An optional string record that specifies the Initials of the author of the presentation comment
         */
        private CString commentRecord;

        /**
         * A Comment2000Atom record that specifies the Settings for displaying the presentation comment
         */
        private Comment2000Atom commentAtom;

        /**
         * Returns the Comment2000Atom of this Comment
         */
        public Comment2000Atom GetComment2000Atom() { return commentAtom; }

        /**
         * Get the Author of this comment
         */
        public String Author
        {
            get
            {
                return authorRecord == null ? null : authorRecord.Text;
            }
            set 
            {
                authorRecord.Text = value;
            }
        }

        /**
         * Get the Author's Initials of this comment
         */
        public String AuthorInitials
        {
            get
            {
                return authorInitialsRecord == null ? null : authorInitialsRecord.Text;
            }
            set 
            {
                authorInitialsRecord.Text = value;
            }
        }

        /**
         * Get the text of this comment
         */
        public String Text
        {
            get
            {
                return commentRecord == null ? null : commentRecord.Text;
            }
            set 
            {
                commentRecord.Text = value;
            }
        }

        /**
         * Set things up, and find our more interesting children
         */
        protected Comment2000(byte[] source, int start, int len)
        {
            // Grab the header
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Find our children
            _children = Record.FindChildRecords(source, start + 8, len - 8);
            FindInterestingChildren();
        }

        /**
         * Go through our child records, picking out the ones that are
         *  interesting, and saving those for use by the easy helper
         *  methods.
         */
        private void FindInterestingChildren()
        {

            foreach (Record r in _children)
            {
                if (r is CString)
                {
                    CString cs = (CString)r;
                    int recInstance = cs.Options >> 4;
                    switch (recInstance)
                    {
                        case 0: authorRecord = cs; break;
                        case 1: commentRecord = cs; break;
                        case 2: authorInitialsRecord = cs; break;
                    }
                }
                else if (r is Comment2000Atom)
                {
                    commentAtom = (Comment2000Atom)r;
                }
            }

        }

        /**
         * Create a new Comment2000, with blank fields
         */
        public Comment2000()
        {
            _header = new byte[8];
            _children = new Record[4];

            // Setup our header block
            _header[0] = 0x0f; // We are a Container record
            LittleEndian.PutShort(_header, 2, (short)_type);

            // Setup our child records
            CString csa = new CString();
            CString csb = new CString();
            CString csc = new CString();
            csa.Options = (0x00);
            csb.Options = (0x10);
            csc.Options = (0x20);
            _children[0] = csa;
            _children[1] = csb;
            _children[2] = csc;
            _children[3] = new Comment2000Atom();
            FindInterestingChildren();
        }

        /**
         * We are of type 1200
         */
        public override long RecordType
        {
            get { return _type; }
        }

        /**
         * Write the contents of the record back, so it can be written
         *  to disk
         */
        public override void WriteOut(Stream out1)
        {
            WriteOut(_header[0], _header[1], _type, _children, out1);
        }

    }
}


