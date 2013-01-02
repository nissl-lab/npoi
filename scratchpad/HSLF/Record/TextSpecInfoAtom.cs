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
    using System.Collections;

    /**
     * The special info Runs Contained in this text.
     * Special info Runs consist of character properties which don?t follow styles.
     *
     * @author Yegor Kozlov
     */
    public class TextSpecInfoAtom : RecordAtom
    {
        /**
         * Record header.
         */
        private byte[] _header;

        /**
         * Record data.
         */
        private byte[] _data;

        /**
         * Constructs the link related atom record from its
         *  source data.
         *
         * @param source the source data as a byte array.
         * @param start the start offset into the byte array.
         * @param len the length of the slice in the byte array.
         */
        protected TextSpecInfoAtom(byte[] source, int start, int len)
        {
            // Get the header.
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Get the record data.
            _data = new byte[len - 8];
            Array.Copy(source, start + 8, _data, 0, len - 8);

        }
        /**
         * Gets the record type.
         * @return the record type.
         */
        public override long RecordType
        {
            get { return RecordTypes.TextSpecInfoAtom.typeID; }
        }

        /**
         * Write the contents of the record back, so it can be written
         * to disk
         *
         * @param out the output stream to write to.
         * @throws java.io.IOException if an error occurs.
         */
        public override void WriteOut(Stream out1)
        {
            out1.Write(_header, (int)out1.Position, _header.Length);
            out1.Write(_data, (int)out1.Position, _data.Length);

        }

        /**
         * Update the text length
         *
         * @param size the text length
         */
        public void SetTextSize(int size)
        {
            LittleEndian.PutInt(_data, 0, size);
        }

        /**
         * Reset the content to one info run with the default values
         * @param size  the site of parent text
         */
        public void Reset(int size)
        {
            _data = new byte[10];
            // 01 00 00 00
            LittleEndian.PutInt(_data, 0, size);
            // 01 00 00 00
            LittleEndian.PutInt(_data, 4, 1); //mask
            // 00 00
            LittleEndian.PutShort(_data, 8, (short)0); //langId

            // Update the size (header bytes 5-8)
            LittleEndian.PutInt(_header, 4, _data.Length);
        }

        /**
         * Get the number of characters covered by this records
         *
         * @return the number of characters covered by this records
         */
        public int GetCharactersCovered()
        {
            int covered = 0;
            TextSpecInfoRun[] Runs = GetTextSpecInfoRuns();
            for (int i = 0; i < Runs.Length; i++) covered += Runs[i].len;
            return covered;
        }

        public TextSpecInfoRun[] GetTextSpecInfoRuns()
        {
            ArrayList lst = new ArrayList();
            int pos = 0;
            int[] bits = { 1, 0, 2 };
            while (pos < _data.Length)
            {
                TextSpecInfoRun run = new TextSpecInfoRun();
                run.len = LittleEndian.GetInt(_data, pos); pos += 4;
                run.mask = LittleEndian.GetInt(_data, pos); pos += 4;
                for (int i = 0; i < bits.Length; i++)
                {
                    if ((run.mask & 1 << bits[i]) != 0)
                    {
                        switch (bits[i])
                        {
                            case 0:
                                run.spellInfo = LittleEndian.GetShort(_data, pos); pos += 2;
                                break;
                            case 1:
                                run.langId = LittleEndian.GetShort(_data, pos); pos += 2;
                                break;
                            case 2:
                                run.altLangId = LittleEndian.GetShort(_data, pos); pos += 2;
                                break;
                        }
                    }
                }
                lst.Add(run);
            }
            return (TextSpecInfoRun[])lst.ToArray(typeof(TextSpecInfoRun));

        }

        public class TextSpecInfoRun
        {
            //Length of special info Run.
            internal int len;

            //Special info mask of this Run;
            internal int mask;

            // info fields as indicated by the mask.
            // -1 means the bit is not set
            internal short spellInfo = -1;
            internal short langId = -1;
            internal short altLangId = -1;

            /**
             * Spelling status of this text. See Spell Info table below.
             *
             * <p>Spell Info Types:</p>
             * <li>0    UnChecked
             * <li>1    Previously incorrect, needs reChecking
             * <li>2    Correct
             * <li>3    Incorrect
             *
             * @return Spelling status of this text
             */
            public short SpellInfo
            {
                get
                {
                    return spellInfo;
                }
            }

            /**
             * Windows LANGID for this text.
             *
             * @return Windows LANGID for this text.
             */
            public short LangId
            {
                get
                {
                    return spellInfo;
                }
            }

            /**
             * Alternate Windows LANGID of this text;
             * must be a valid non-East Asian LANGID if the text has an East Asian language,
             * otherwise may be an East Asian LANGID or language neutral (zero).
             *
             * @return  Alternate Windows LANGID of this text
             */
            public short AltLangId
            {
                get
                {
                    return altLangId;
                }
            }

            /**
             * @return Length of special info Run.
             */
            public int Length
            {
                get
                {
                    return len;
                }
            }
        }
    }
}