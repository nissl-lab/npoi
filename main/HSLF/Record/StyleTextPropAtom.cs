/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using NPOI.HSLF.Exceptions;
using NPOI.HSLF.Model;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static NPOI.HSLF.Model.TextPropCollection;

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
 */
	public class StyleTextPropAtom: RecordAtom
	{
		public static  long _type = RecordTypes.StyleTextPropAtom.typeID;

		private  byte[] _header;
		private byte[] reserved;

		private byte[] rawContents; // Holds the contents between write-outs

		/**
		 * Only set to true once setParentTextSize(int) is called.
		 * Until then, no stylings will have been decoded
		 */
		private bool initialised;

		/**
		 * The list of all the different paragraph stylings we code for.
		 * Each entry is a TextPropCollection, which tells you how many
		 *  Characters the paragraph covers, and also contains the TextProps
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
		 *  Characters the character styling covers, and also contains the
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
			//return styles.stream().mapToInt(TextPropCollection::getCharactersCovered).sum();
			return styles.Select(s => s.GetCharactersCovered()).Sum();
		}

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
					throw new HSLFException("Not enough data to form a StyleTextPropAtom (min size 18 bytes long) - found " + (source.length - start));
				}
			}

			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Save the contents of the atom, until we're asked to go and
			//  decode them (via a call to setParentTextSize(int)
			rawContents = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());
			reserved = new byte[0];

			// Set empty lists, ready for when they call setParentTextSize
			paragraphStyles = new List<TextPropCollection>();
			charStyles = new List<TextPropCollection>();
		}


		/**
		 * A new set of text style properties for some text without any.
		 */
		public StyleTextPropAtom(int parentTextSize)
		{
			_header = new byte[8];
			rawContents = new byte[0];
			reserved = new byte[0];

			// Set our type
			LittleEndian.PutInt(_header, 2, (short)_type);
			// Our initial size is 10
			LittleEndian.PutInt(_header, 4, 10);

			// Set empty paragraph and character styles
			paragraphStyles = new List<TextPropCollection>();
			charStyles = new List<TextPropCollection>();

			AddParagraphTextPropCollection(parentTextSize);
			AddCharacterTextPropCollection(parentTextSize);

			// Set us as now initialised
			initialised = true;

			try
			{
				UpdateRawContents();
			}
			catch (IOException e)
			{
				throw new HSLFException(e);
			}
		}


		/**
		 * We are of type 4001
		 */
		//@Override
		public override long GetRecordType() { return _type; }


		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		//@Override
		public override void WriteOut(OutputStream _out)
		{
			// First thing to do is update the raw bytes of the contents, based
			//  on the properties
			UpdateRawContents();

			// Write out the (new) header
			_out.Write(_header);

			// Write out the styles
			_out.Write(rawContents);

			// Write out any extra bits
			_out.Write(reserved);
		}


		/**
		 * Tell us how much text the parent TextCharsAtom or TextBytesAtom
		 *  contains, so we can go ahead and initialise ourselves.
		 */
		public void SetParentTextSize(int size)
		{
			if (initialised)
			{
				return;
			}

			int pos = 0;
			int textHandled = 0;

			paragraphStyles.Clear();
			charStyles.Clear();

			// While we have text in need of paragraph stylings, go ahead and
			// grok the contents as paragraph formatting data
			int prsize = size;
			while (pos < rawContents.Length && textHandled < prsize)
			{
				// First up, fetch the number of characters this applies to
				int textLen = LittleEndian.GetInt(rawContents, pos);
				textLen = checkTextLength(textLen, textHandled, size);
				textHandled += textLen;
				pos += 4;

				short indent = LittleEndian.GetShort(rawContents, pos);
				pos += 2;

				// Grab the 4 byte value that tells us what properties follow
				int paraFlags = LittleEndian.GetInt(rawContents, pos);
				pos += 4;

				// Now make sense of those properties
				TextPropCollection thisCollection = new TextPropCollection(textLen, TextPropType.paragraph);
				thisCollection.setIndentLevel(indent);
				int plSize = thisCollection.BuildTextPropList(paraFlags, rawContents, pos);
				pos += plSize;

				// Save this properties set
				paragraphStyles.Add(thisCollection);

				// Handle extra 1 paragraph styles at the end
				if (pos < rawContents.Length && textHandled == size)
				{
					prsize++;
				}

			}
			if (rawContents.Length > 0 && textHandled != (size + 1))
			{
				//LOG.atWarn().log("Problem reading paragraph style runs: textHandled = {}, text.size+1 = {}", box(textHandled), box(size + 1));
			}

			// Now do the character stylings
			textHandled = 0;
			int chsize = size;
			while (pos < rawContents.Length && textHandled < chsize)
			{
				// First up, fetch the number of characters this applies to
				int textLen = LittleEndian.GetInt(rawContents, pos);
				textLen = checkTextLength(textLen, textHandled, size);
				textHandled += textLen;
				pos += 4;

				// Grab the 4 byte value that tells us what properties follow
				int charFlags = LittleEndian.GetInt(rawContents, pos);
				pos += 4;

				// Now make sense of those properties
				// (Assuming we actually have some)
				TextPropCollection thisCollection = new TextPropCollection(textLen, TextPropType.character);
				int chSize = thisCollection.BuildTextPropList(charFlags, rawContents, pos);
				pos += chSize;

				// Save this properties set
				charStyles.Add(thisCollection);

				// Handle extra 1 char styles at the end
				if (pos < rawContents.Length && textHandled == size)
				{
					chsize++;
				}
			}
			if (rawContents.Length > 0 && textHandled != (size + 1))
			{
				//LOG.atWarn().log("Problem reading character style runs: textHandled = {}, text.size+1 = {}", box(textHandled), box(size + 1));
			}

			// Handle anything left over
			if (pos < rawContents.Length)
			{
				reserved = IOUtils.SafelyClone(rawContents, pos, rawContents.Length - pos, rawContents.Length);
			}

			initialised = true;
		}

		private int CheckTextLength(int readLength, int handledSoFar, int overallSize)
		{
			if (readLength + handledSoFar > overallSize + 1)
			{
				//LOG.atWarn().log("Style length of {} at {} larger than stated size of {}, truncating", box(readLength), box(handledSoFar), box(overallSize));
				return overallSize + 1 - handledSoFar;
			}
			return readLength;
		}


		/**
		 * Updates the cache of the raw contents. Serialised the styles out.
		 */
		private void UpdateRawContents()
		{
			if (initialised) {
				// Only update the style bytes, if the styles have been potentially changed
				using (MemoryStream baos = new MemoryStream()) {
					// First up, we need to serialise the paragraph properties
					foreach (TextPropCollection tpc in paragraphStyles)
					{
						tpc.WriteOut((OutputStream)baos);
					}

					// Now, we do the character ones
					foreach (TextPropCollection tpc in charStyles)
					{
						tpc.WriteOut((OutputStream)baos);
					}

					rawContents = baos.ToArray();
				}
			}

			// Now ensure that the header size is correct
			int newSize = rawContents.Length + reserved.Length;
			LittleEndian.PutInt(_header, 4, newSize);
		}

		/**
		 * Clear styles, so new collections can be added
		 */
		public void ClearStyles()
		{
			paragraphStyles.Clear();
			charStyles.Clear();
			reserved = new byte[0];
			initialised = true;
		}

		/**
		 * Create a new Paragraph TextPropCollection, and add it to the list
		 * @param charactersCovered The number of characters this TextPropCollection will cover
		 * @return the new TextPropCollection, which will then be in the list
		 */
		public TextPropCollection AddParagraphTextPropCollection(int charactersCovered)
		{
			TextPropCollection tpc = new TextPropCollection(charactersCovered, TextPropType.paragraph);
			paragraphStyles.Add(tpc);
			return tpc;
		}

		public void AddParagraphTextPropCollection(TextPropCollection tpc)
		{
			paragraphStyles.Add(tpc);
		}

		/**
		 * Create a new Character TextPropCollection, and add it to the list
		 * @param charactersCovered The number of characters this TextPropCollection will cover
		 * @return the new TextPropCollection, which will then be in the list
		 */
		public TextPropCollection AddCharacterTextPropCollection(int charactersCovered)
		{
			TextPropCollection tpc = new TextPropCollection(charactersCovered, TextPropType.character);
			charStyles.Add(tpc);
			return tpc;
		}

		public void AddCharacterTextPropCollection(TextPropCollection tpc)
		{
			charStyles.Add(tpc);
		}

		/* ************************************************************************ */


		/**
		 * @return the string representation of the record data
		 */
		//@Override
		public override string ToString()
		{
			StringBuilder _out = new StringBuilder();

			_out.Append("StyleTextPropAtom:\n");
			if (!initialised)
			{
				_out.Append("Uninitialised, dumping Raw Style Data\n");
			}
			else
			{

				_out.Append("Paragraph properties\n");
				foreach (TextPropCollection pr in GetParagraphStyles())
				{
					_out.Append(pr);
				}

				_out.Append("Character properties\n");
				foreach (TextPropCollection pr in GetCharacterStyles())
				{
					_out.Append(pr);
				}

				_out.Append("Reserved bytes\n");
				_out.Append(HexDump.Dump(reserved, 0, 0));
			}

			_out.Append("  original byte stream \n");

			byte[] buf = IOUtils.SafelyAllocate(rawContents.Length + (long)reserved.Length, GetMaxRecordLength());
			Array.Copy(rawContents, 0, buf, 0, rawContents.Length);
			Array.Copy(reserved, 0, buf, rawContents.Length, reserved.Length);
			_out.Append(HexDump.Dump(buf, 0, 0));

			return _out.ToString();
		}

	//@Override
	public override IDictionary<string, Func<T>> GetGenericProperties<T>()
	{
		return (IDictionary<string, Func<T>>)(!initialised ? null : GenericRecordUtil.GetGenericProperties(
			"paragraphStyles", GetParagraphStyles,
			"characterStyles", GetCharacterStyles
		));
	}
}
}
