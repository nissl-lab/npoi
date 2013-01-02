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

using NPOI.Util;
using System.Collections.Generic;
using System.IO;
using System;
using NPOI.HSLF.Record;
namespace NPOI.HSLF.Model.TextProperties
{
    /**
     * For a given run of characters, holds the properties (which could
     *  be paragraph properties or character properties).
     * Used to hold the number of characters affected, the list of active
     *  properties, and the random reserved field if required.
     */
    public class TextPropCollection
    {
        private int charactersCovered;
        private short reservedField;
        private List<TextProp> textPropList;
        private int maskSpecial = 0;

        public int GetSpecialMask() { return maskSpecial; }

        /** Fetch the number of characters this styling applies to */
        public int GetCharactersCovered() { return charactersCovered; }
        /** Fetch the TextProps that define this styling */
        public List<TextProp> GetTextPropList() { return textPropList; }

        /** Fetch the TextProp with this name, or null if it isn't present */
        public TextProp FindByName(String textPropName)
        {
            for (int i = 0; i < textPropList.Count; i++)
            {
                TextProp prop = textPropList[i];
                if (prop.GetName().Equals(textPropName))
                {
                    return prop;
                }
            }
            return null;
        }

        /** Add the TextProp with this name to the list */
        public TextProp AddWithName(String name) {
		// Find the base TextProp to base on
		TextProp base1 = null;
		for(int i=0; i < StyleTextPropAtom.characterTextPropTypes.Length; i++) {
			if(StyleTextPropAtom.characterTextPropTypes[i].GetName().Equals(name)) {
				base1 = StyleTextPropAtom.characterTextPropTypes[i];
			}
		}
		for(int i=0; i < StyleTextPropAtom.paragraphTextPropTypes.Length; i++) {
			if(StyleTextPropAtom.paragraphTextPropTypes[i].GetName().Equals(name)) {
				base1 = StyleTextPropAtom.paragraphTextPropTypes[i];
			}
		}
		if(base1== null) {
			throw new ArgumentException("No TextProp with name " + name + " is defined to add from");
		}
		
		// Add a copy of this property, in the right place to the list
		TextProp textProp = (TextProp)base1.Clone();
		int pos = 0;
		for(int i=0; i<textPropList.Count; i++) {
			TextProp curProp = textPropList[i];
			if(textProp.GetMask() > curProp.GetMask()) {
				pos++;
			}
		}
		textPropList.Insert(pos, textProp);
		return textProp;
	}

        /**
         * For an existing Set of text properties, build the list of 
         *  properties coded for in a given run of properties.
         * @return the number of bytes that were used encoding the properties list
         */
        public int BuildTextPropList(int ContainsField, TextProp[] potentialProperties, byte[] data, int dataOffset)
        {
            int bytesPassed = 0;

            // For each possible entry, see if we match the mask
            // If we do, decode that, save it, and shuffle on
            for (int i = 0; i < potentialProperties.Length; i++)
            {
                // Check there's still data left to read

                // Check if this property is found in the mask
                if ((ContainsField & potentialProperties[i].GetMask()) != 0)
                {
                    if (dataOffset + bytesPassed >= data.Length)
                    {
                        // Out of data, can't be any more properties to go
                        // remember the mask and return
                        maskSpecial |= potentialProperties[i].GetMask();
                        return bytesPassed;
                    }

                    // Bingo, data Contains this property
                    TextProp prop = (TextProp)potentialProperties[i].Clone();
                    int val = 0;
                    if (prop.GetSize() == 2)
                    {
                        val = LittleEndian.GetShort(data, dataOffset + bytesPassed);
                    }
                    else if (prop.GetSize() == 4)
                    {
                        val = LittleEndian.GetInt(data, dataOffset + bytesPassed);
                    }
                    else if (prop.GetSize() == 0)
                    {
                        //remember "special" bits.
                        maskSpecial |= potentialProperties[i].GetMask();
                        continue;
                    }
                    prop.SetValue(val);
                    bytesPassed += prop.GetSize();
                    textPropList.Add(prop);
                }
            }

            // Return how many bytes were used
            return bytesPassed;
        }

        /**
         * Create a new collection of text properties (be they paragraph
         *  or character) which will be groked via a subsequent call to
         *  buildTextPropList().
         */
        public TextPropCollection(int charactersCovered, short reservedField)
        {
            this.charactersCovered = charactersCovered;
            this.reservedField = reservedField;
            textPropList = new List<TextProp>();
        }

        /**
         * Create a new collection of text properties (be they paragraph
         *  or character) for a run of text without any
         */
        public TextPropCollection(int textSize)
        {
            charactersCovered = textSize;
            reservedField = -1;
            textPropList = new List<TextProp>();
        }

        /**
         * Update the size of the text that this Set of properties
         *  applies to 
         */
        public void updateTextSize(int textSize)
        {
            charactersCovered = textSize;
        }

        /**
         * Writes out to disk the header, and then all the properties
         */
        public void WriteOut(Stream o)
        {
            // First goes the number of characters we affect
            StyleTextPropAtom.WriteLittleEndian(charactersCovered, o);

            // Then we have the reserved field if required
            if (reservedField > -1)
            {
                StyleTextPropAtom.WriteLittleEndian(reservedField, o);
            }

            // Then the mask field
            int mask = maskSpecial;
            for (int i = 0; i < textPropList.Count; i++)
            {
                TextProp textProp = (TextProp)textPropList[i];
                //sometimes header indicates that the bitmask is present but its value is 0

                if (textProp is BitMaskTextProp)
                {
                    if (mask == 0) mask |= textProp.GetWriteMask();
                }
                else
                {
                    mask |= textProp.GetWriteMask();
                }
            }
            StyleTextPropAtom.WriteLittleEndian(mask, o);

            // Then the contents of all the properties
            for (int i = 0; i < textPropList.Count; i++)
            {
                TextProp textProp = textPropList[i];
                int val = textProp.GetValue();
                if (textProp.GetSize() == 2)
                {
                    StyleTextPropAtom.WriteLittleEndian((short)val, o);
                }
                else if (textProp.GetSize() == 4)
                {
                    StyleTextPropAtom.WriteLittleEndian(val, o);
                }
            }
        }

        public short GetReservedField()
        {
            return reservedField;
        }

        public void SetReservedField(short val)
        {
            reservedField = val;
        }
    }


}