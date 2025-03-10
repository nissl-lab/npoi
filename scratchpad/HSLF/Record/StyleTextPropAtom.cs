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


    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using NPOI.HSLF.Record;
using NPOI.Util;
using NPOI.HSLF.Model.TextProperties;
    
    namespace NPOI.HSLF.Record
    {

        /**
         * A StyleTextPropAtom (type 4001). Holds basic character properties
         *  (bold, italic, underline, font size etc) and paragraph properties
         *  (alignment, line spacing etc) for the block of text (TextBytesAtom
         *  or TextCharsAtom) that this record follows.
         * You will find two lists within this class.
         *  1 - Paragraph style list (paragraphStyles)
         *  2 - Character style list (charStyles)
         * Both are lists of TextPropCollections. These define how many characters
         *  the style applies to, and what style elements make up the style (another
         *  list, this time of TextProps). Each TextProp has a value, which somehow
         *  encapsulates a property of the style
         *
         * @author Nick Burch
         * @author Yegor Kozlov
         */

        public class StyleTextPropAtom : RecordAtom
        {
            private byte[] _header;
            private static long _type = 4001L;
            private byte[] reserved;

            private byte[] rawContents; // Holds the contents between Write-outs

            /**
             * Only Set to true once SetParentTextSize(int) is called.
             * Until then, no stylings will have been decoded
             */
            private bool Initialised = false;

            /**
             * The list of all the different paragraph stylings we code for.
             * Each entry is a TextPropCollection, which tells you how many
             *  Characters the paragraph covers, and also Contains the TextProps
             *  that actually define the styling of the paragraph.
             */
            private List<TextPropCollection> paragraphStyles;
            public List<TextPropCollection> GetParagraphStyles() { return paragraphStyles; }
            /**
             * Updates the link list of TextPropCollections which make up the
             *  paragraph stylings
             */
            public void SetParagraphStyles(List<TextPropCollection> ps) { paragraphStyles = ps; }
            /**
             * The list of all the different character stylings we code for.
             * Each entry is a TextPropCollection, which tells you how many
             *  Characters the character styling covers, and also Contains the
             *  TextProps that actually define the styling of the characters.
             */
            private List<TextPropCollection> charStyles;
            public List<TextPropCollection> GetCharacterStyles() { return charStyles; }
            /**
             * Updates the link list of TextPropCollections which make up the
             *  character stylings
             */
            public void SetCharacterStyles(List<TextPropCollection> cs) { charStyles = cs; }

            /**
             * Returns how many characters the paragraph's
             *  TextPropCollections cover.
             * (May be one or two more than the underlying text does,
             *  due to having extra characters meaning something
             *  special to powerpoint)
             */
            public int GetParagraphTextLengthCovered()
            {
                return GetCharactersCovered(paragraphStyles);
            }
            /**
             * Returns how many characters the character's
             *  TextPropCollections cover.
             * (May be one or two more than the underlying text does,
             *  due to having extra characters meaning something
             *  special to powerpoint)
             */
            public int GetCharacterTextLengthCovered()
            {
                return GetCharactersCovered(charStyles);
            }
            private int GetCharactersCovered(List<TextPropCollection> styles)
            {
                int length = 0;
                foreach (TextPropCollection tpc in styles)
                {
                    length += tpc.GetCharactersCovered();
                }
                return length;
            }

            /** All the different kinds of paragraph properties we might handle */
            public static TextProp[] paragraphTextPropTypes = new TextProp[] {
                new TextProp(0, 0x1, "hasBullet"),
                new TextProp(0, 0x2, "hasBulletFont"),
                new TextProp(0, 0x4, "hasBulletColor"),
                new TextProp(0, 0x8, "hasBulletSize"),
                new ParagraphFlagsTextProp(),
                new TextProp(2, 0x80, "bullet.char"),
				new TextProp(2, 0x10, "bullet.font"),
                new TextProp(2, 0x40, "bullet.size"),
				new TextProp(4, 0x20, "bullet.color"),
                new AlignmentTextProp(),
                new TextProp(2, 0x100, "text.offset"),
                new TextProp(2, 0x400, "bullet.offset"),
                new TextProp(2, 0x1000, "linespacing"),
                new TextProp(2, 0x2000, "spaceBefore"),
                new TextProp(2, 0x4000, "spaceAfter"),
                new TextProp(2, 0x8000, "defaultTabSize"),
				new TextProp(2, 0x100000, "tabStops"),
				new TextProp(2, 0x10000, "fontAlign"),
				new TextProp(2, 0xA0000, "wrapFlags"),
				new TextProp(2, 0x200000, "textDirection")
	};
            /** All the different kinds of character properties we might handle */
            public static TextProp[] characterTextPropTypes = new TextProp[] {
                new TextProp(0, 0x1, "bold"),
                new TextProp(0, 0x2, "italic"),
                new TextProp(0, 0x4, "underline"),
                new TextProp(0, 0x8, "unused1"),
                new TextProp(0, 0x10, "shadow"),
                new TextProp(0, 0x20, "fehint"),
                new TextProp(0, 0x40, "unused2"),
                new TextProp(0, 0x80, "kumi"),
                new TextProp(0, 0x100, "unused3"),
                new TextProp(0, 0x200, "emboss"),
                new TextProp(0, 0x400, "nibble1"),
                new TextProp(0, 0x800, "nibble2"),
                new TextProp(0, 0x1000, "nibble3"),
                new TextProp(0, 0x2000, "nibble4"),
                new TextProp(0, 0x4000, "unused4"),
                new TextProp(0, 0x8000, "unused5"),
                new CharFlagsTextProp(),
				new TextProp(2, 0x10000, "font.index"),
                new TextProp(0, 0x100000, "pp10ext"),
                new TextProp(2, 0x200000, "asian.font.index"),
                new TextProp(2, 0x400000, "ansi.font.index"),
                new TextProp(2, 0x800000, "symbol.font.index"),
				new TextProp(2, 0x20000, "font.size"),
				new TextProp(4, 0x40000, "font.color"),
                new TextProp(2, 0x80000, "superscript"),

    };

            /* *************** record code follows ********************** */

            /**
             * For the Text Style Properties (StyleTextProp) Atom
             */
            public StyleTextPropAtom(byte[] source, int start, int len)
            {
                // Sanity Checking - we're always at least 8+10 bytes long
                if (len < 18)
                {
                    len = 18;
                    if (source.Length - start < 18)
                    {
                        throw new Exception("Not enough data to form a StyleTextPropAtom (min size 18 bytes long) - found " + (source.Length - start));
                    }
                }

                // Get the header
                _header = new byte[8];
                Array.Copy(source, start, _header, 0, 8);

                // Save the contents of the atom, until we're asked to go and
                //  decode them (via a call to SetParentTextSize(int)
                rawContents = new byte[len - 8];
                Array.Copy(source, start + 8, rawContents, 0, rawContents.Length);
                reserved = Array.Empty<byte>();

                // Set empty linked lists, Ready for when they call SetParentTextSize
                paragraphStyles = new List<TextPropCollection>();
                charStyles = new List<TextPropCollection>();
            }


            /**
             * A new Set of text style properties for some text without any.
             */
            public StyleTextPropAtom(int parentTextSize)
            {
                _header = new byte[8];
                rawContents = Array.Empty<byte>();
                reserved = Array.Empty<byte>();

                // Set our type
                LittleEndian.PutInt(_header, 2, (short)_type);
                // Our Initial size is 10
                LittleEndian.PutInt(_header, 4, 10);

                // Set empty paragraph and character styles
                paragraphStyles = new List<TextPropCollection>();
                charStyles = new List<TextPropCollection>();

                TextPropCollection defaultParagraphTextProps =
                    new TextPropCollection(parentTextSize, (short)0);
                paragraphStyles.Add(defaultParagraphTextProps);

                TextPropCollection defaultCharacterTextProps =
                    new TextPropCollection(parentTextSize);
                charStyles.Add(defaultCharacterTextProps);

                // Set us as now Initialised
                Initialised = true;
            }


            /**
             * We are of type 4001
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
                // First thing to do is update the raw bytes of the contents, based
                //  on the properties
                updateRawContents();

                // Now ensure that the header size is correct
                int newSize = rawContents.Length + reserved.Length;
                LittleEndian.PutInt(_header, 4, newSize);

                // Write out the (new) header
                out1.Write(_header,0,_header.Length);

                // Write out the styles
                out1.Write(rawContents,0,rawContents.Length);

                // Write out any extra bits
                out1.Write(reserved,0,reserved.Length);
            }


            /**
             * Tell us how much text the parent TextCharsAtom or TextBytesAtom
             *  Contains, so we can go ahead and Initialise ourselves.
             */
            public void SetParentTextSize(int size)
            {
                int pos = 0;
                int textHandled = 0;

                // While we have text in need of paragraph stylings, go ahead and
                // grok the contents as paragraph formatting data
                int prsize = size;
                while (pos < rawContents.Length && textHandled < prsize)
                {
                    // First up, fetch the number of characters this applies to
                    int textLen = LittleEndian.GetInt(rawContents, pos);
                    textHandled += textLen;
                    pos += 4;

                    short indent = LittleEndian.GetShort(rawContents, pos);
                    pos += 2;

                    // Grab the 4 byte value that tells us what properties follow
                    int paraFlags = LittleEndian.GetInt(rawContents, pos);
                    pos += 4;

                    // Now make sense of those properties
                    TextPropCollection thisCollection = new TextPropCollection(textLen, indent);
                    int plSize = thisCollection.BuildTextPropList(
                            paraFlags, paragraphTextPropTypes, rawContents, pos);
                    pos += plSize;

                    // Save this properties Set
                    paragraphStyles.Add(thisCollection);

                    // Handle extra 1 paragraph styles at the end
                    if (pos < rawContents.Length && textHandled == size)
                    {
                        prsize++;
                    }

                }
                if (rawContents.Length > 0 && textHandled != (size + 1))
                {
                    logger.Log(POILogger.WARN, "Problem Reading paragraph style Runs: textHandled = " + textHandled + ", text.size+1 = " + (size + 1));
                }

                // Now do the character stylings
                textHandled = 0;
                int chsize = size;
                while (pos < rawContents.Length && textHandled < chsize)
                {
                    // First up, fetch the number of characters this applies to
                    int textLen = LittleEndian.GetInt(rawContents, pos);
                    textHandled += textLen;
                    pos += 4;

                    // There is no 2 byte value
                    short no_val = -1;

                    // Grab the 4 byte value that tells us what properties follow
                    int charFlags = LittleEndian.GetInt(rawContents, pos);
                    pos += 4;

                    // Now make sense of those properties
                    // (Assuming we actually have some)
                    TextPropCollection thisCollection = new TextPropCollection(textLen, no_val);
                    int chSize = thisCollection.BuildTextPropList(
                            charFlags, characterTextPropTypes, rawContents, pos);
                    pos += chSize;

                    // Save this properties Set
                    charStyles.Add(thisCollection);

                    // Handle extra 1 char styles at the end
                    if (pos < rawContents.Length && textHandled == size)
                    {
                        chsize++;
                    }
                }
                if (rawContents.Length > 0 && textHandled != (size + 1))
                {
                    logger.Log(POILogger.WARN, "Problem Reading character style Runs: textHandled = " + textHandled + ", text.size+1 = " + (size + 1));
                }

                // Handle anything left over
                if (pos < rawContents.Length)
                {
                    reserved = new byte[rawContents.Length - pos];
                    Array.Copy(rawContents, pos, reserved, 0, reserved.Length);
                }

                Initialised = true;
            }


            /**
             * Updates the cache of the raw contents. Serialised the styles out.
             */
            private void updateRawContents()
            {
                if (!Initialised)
                {
                    // We haven't groked the styles since creation, so just stick
                    // with what we found
                    return;
                }

                MemoryStream baos = new MemoryStream();

                // First up, we need to serialise the paragraph properties
                for (int i = 0; i < paragraphStyles.Count; i++)
                {
                    TextPropCollection tpc = paragraphStyles[i];
                    tpc.WriteOut(baos);
                }

                // Now, we do the character ones
                for (int i = 0; i < charStyles.Count; i++)
                {
                    TextPropCollection tpc = charStyles[i];
                    tpc.WriteOut(baos);
                }

                rawContents = baos.ToArray();
            }

            public void SetRawContents(byte[] bytes)
            {
                rawContents = bytes;
                reserved = Array.Empty<byte>();
                Initialised = false;
            }

            /**
             * Create a new Paragraph TextPropCollection, and add it to the list
             * @param charactersCovered The number of characters this TextPropCollection will cover
             * @return the new TextPropCollection, which will then be in the list
             */
            public TextPropCollection AddParagraphTextPropCollection(int charactersCovered)
            {
                TextPropCollection tpc = new TextPropCollection(charactersCovered, (short)0);
                paragraphStyles.Add(tpc);
                return tpc;
            }
            /**
             * Create a new Character TextPropCollection, and add it to the list
             * @param charactersCovered The number of characters this TextPropCollection will cover
             * @return the new TextPropCollection, which will then be in the list
             */
            public TextPropCollection AddCharacterTextPropCollection(int charactersCovered)
            {
                TextPropCollection tpc = new TextPropCollection(charactersCovered);
                charStyles.Add(tpc);
                return tpc;
            }

            /* ************************************************************************ */


            /**
             * Dump the record content into <code>StringBuilder</code>
             *
             * @return the string representation of the record data
             */
            public override String ToString()
            {
                StringBuilder out1 = new StringBuilder();

                out1.Append("StyleTextPropAtom:\n");
                if (!Initialised)
                {
                    out1.Append("Uninitialised, dumping Raw Style Data\n");
                }
                else
                {

                    out1.Append("Paragraph properties\n");

                    foreach (TextPropCollection pr in GetParagraphStyles())
                    {
                        out1.Append("  chars covered: " + pr.GetCharactersCovered());
                        out1.Append("  special mask flags: 0x" + HexDump.ToHex(pr.GetSpecialMask()) + "\n");
                        foreach (TextProp p in pr.GetTextPropList())
                        {
                            out1.Append("    " + p.GetName() + " = " + p.GetValue());
                            out1.Append(" (0x" + HexDump.ToHex(p.GetValue()) + ")\n");
                        }

                        out1.Append("  para bytes that would be written: \n");

                        MemoryStream baos = new MemoryStream();
                        pr.WriteOut(baos);
                        byte[] b = baos.ToArray();
                        out1.Append(HexDump.Dump(b, 0, 0));

                    }

                    out1.Append("Character properties\n");
                    foreach (TextPropCollection pr in GetCharacterStyles())
                    {
                        out1.Append("  chars covered: " + pr.GetCharactersCovered());
                        out1.Append("  special mask flags: 0x" + HexDump.ToHex(pr.GetSpecialMask()) + "\n");
                        foreach (TextProp p in pr.GetTextPropList())
                        {
                            out1.Append("    " + p.GetName() + " = " + p.GetValue());
                            out1.Append(" (0x" + HexDump.ToHex(p.GetValue()) + ")\n");
                        }

                        out1.Append("  char bytes that would be written: \n");

                        MemoryStream baos = new MemoryStream();
                        pr.WriteOut(baos);
                        byte[] b = baos.ToArray();
                        out1.Append(HexDump.Dump(b, 0, 0));

                    }
                }

                out1.Append("  original byte stream \n");
                out1.Append(HexDump.Dump(rawContents, 0, 0));

                return out1.ToString();
            }
        }

    }