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
     * Ruler of a text as it differs from the style's ruler Settings.
     *
     * @author Yegor Kozlov
     */
    public class TextRulerAtom : RecordAtom
    {

        /**
         * Record header.
         */
        private byte[] _header;

        /**
         * Record data.
         */
        private byte[] _data;

        //ruler internals
        private int defaultTabSize;
        private int numLevels;
        private int[] tabStops;
        private int[] bulletOffSets = new int[5];
        private int[] textOffSets = new int[5];

        /**
         * Constructs a new empty ruler atom.
         */
        public TextRulerAtom()
        {
            _header = new byte[8];
            _data = Array.Empty<byte>();

            LittleEndian.PutShort(_header, 2, (short)RecordType);
            LittleEndian.PutInt(_header, 4, _data.Length);
        }

        /**
         * Constructs the ruler atom record from its
         *  source data.
         *
         * @param source the source data as a byte array.
         * @param start the start offset into the byte array.
         * @param len the length of the slice in the byte array.
         */
        protected TextRulerAtom(byte[] source, int start, int len)
        {
            // Get the header.
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Get the record data.
            _data = new byte[len - 8];
            Array.Copy(source, start + 8, _data, 0, len - 8);

            Read();
        }

        /**
         * Gets the record type.
         *
         * @return the record type.
         */
        public override long RecordType
        {
            get
            {
                return RecordTypes.TextRulerAtom.typeID;
            }
        }

        /**
         * Write the contents of the record back, so it can be written
         * to disk.
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
         * Read the record bytes and Initialize the internal variables
         */
        private void Read()
        {
            int pos = 0;
            short mask = LittleEndian.GetShort(_data); pos += 4;
            short val;
            int[] bits = { 1, 0, 2, 3, 8, 4, 9, 5, 10, 6, 11, 7, 12 };
            for (int i = 0; i < bits.Length; i++)
            {
                if ((mask & 1 << bits[i]) != 0)
                {
                    switch (bits[i])
                    {
                        case 0:
                            //defaultTabSize
                            defaultTabSize = LittleEndian.GetShort(_data, pos); pos += 2;
                            break;
                        case 1:
                            //numLevels
                            numLevels = LittleEndian.GetShort(_data, pos); pos += 2;
                            break;
                        case 2:
                            //tabStops
                            val = LittleEndian.GetShort(_data, pos); pos += 2;
                            tabStops = new int[val * 2];
                            for (int j = 0; j < tabStops.Length; j++)
                            {
                                tabStops[j] = LittleEndian.GetUShort(_data, pos); pos += 2;
                            }
                            break;
                        case 3:
                        case 4:
                        case 5:
                        case 6:
                        case 7:
                            //bullet.offset
                            val = LittleEndian.GetShort(_data, pos); pos += 2;
                            bulletOffSets[bits[i] - 3] = val;
                            break;
                        case 8:
                        case 9:
                        case 10:
                        case 11:
                        case 12:
                            //text.offset
                            val = LittleEndian.GetShort(_data, pos); pos += 2;
                            textOffSets[bits[i] - 8] = val;
                            break;
                    }
                }
            }
        }

        /**
         * Default distance between tab stops, in master coordinates (576 dpi).
         */
        public int GetDefaultTabSize()
        {
            return defaultTabSize;
        }

        /**
         * Number of indent levels (maximum 5).
         */
        public int GetNumberOfLevels()
        {
            return numLevels;
        }

        /**
         * Default distance between tab stops, in master coordinates (576 dpi).
         */
        public int[] GetTabStops()
        {
            return tabStops;
        }

        /**
         * Paragraph's distance from shape's left margin, in master coordinates (576 dpi).
         */
        public int[] GetTextOffSets()
        {
            return textOffSets;
        }

        /**
         * First line of paragraph's distance from shape's left margin, in master coordinates (576 dpi).
         */
        public int[] GetBulletOffSets()
        {
            return bulletOffSets;
        }

        public static TextRulerAtom GetParagraphInstance()
        {
            byte[] data = new byte[] {
            0x00, 0x00, (byte)0xA6, 0x0F, 0x0A, 0x00, 0x00, 0x00,
            0x10, 0x03, 0x00, 0x00, (byte)0xF9, 0x00, 0x41, 0x01, 0x41, 0x01
        };
            TextRulerAtom ruler = new TextRulerAtom(data, 0, data.Length);
            return ruler;
        }

        public void SetParagraphIndent(short tetxOffSet, short bulletOffSet)
        {
            LittleEndian.PutShort(_data, 4, tetxOffSet);
            LittleEndian.PutShort(_data, 6, bulletOffSet);
            LittleEndian.PutShort(_data, 8, bulletOffSet);
        }
    }
}

