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

using NPOI.Common.UserModel.Fonts;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.HSLF.Record
{
	/**
 * {@code FontCollection} ia a container that holds information
 * about all the fonts in the presentation.
 */
	public class FontCollection : RecordContainer
	{
		private Dictionary<int, HSLFFontInfo> fonts = new Dictionary<int, HSLFFontInfo>();
		private byte[] _header;

		/* package */
		FontCollection(byte[] source, int start, int len)
		{
			_header = Arrays.CopyOfRange(source, start, start + 8);

			_children = Record.FindChildRecords(source, start + 8, len - 8);

			foreach (Record r in _children)
			{
				if (r is FontEntityAtom)
				{
					HSLFFontInfo fi = new HSLFFontInfo((FontEntityAtom)r);
					fonts.Add(fi.GetIndex(), fi);
				}
				else if (r is FontEmbeddedData)
				{
					FontEmbeddedData fed = (FontEmbeddedData)r;
					FontHeader fontHeader = fed.GetFontHeader();
					HSLFFontInfo fi = AddFont(fontHeader);
					fi.AddFacet(fed);
				}
				else
				{
					//LOG.atWarn().log("FontCollection child wasn't a FontEntityAtom, was {}", r.getClass().getSimpleName());
				}
			}
		}

		/**
		 * Return the type, which is 2005
		 */
		//@Override
		public override long GetRecordType()
		{
			return RecordTypes.FontCollection.typeID;
		}

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		//@Override
		public override void WriteOut(OutputStream _out)
		{
			WriteOut(_header[0], _header[1], GetRecordType(), _children, _out);
		}

		/**
		 * Add font with the given FontInfo configuration to the font collection.
		 * The returned FontInfo contains the HSLF specific details and the collection
		 * uniquely contains fonts based on their typeface, i.e. calling the method with FontInfo
		 * objects having the same name results in the same HSLFFontInfo reference.
		 *
		 * @param fontInfo the FontInfo configuration, can be a instance of {@link HSLFFontInfo},
		 *      {@link HSLFFontInfoPredefined} or a custom implementation
		 * @return the register HSLFFontInfo object
		 */
		public HSLFFontInfo AddFont(FontInfo fontInfo)
		{
			HSLFFontInfo fi = GetFontInfo(fontInfo.GetTypeface(), fontInfo.GetCharset());
			if (fi != null)
			{
				return fi;
			}

			fi = new HSLFFontInfo(fontInfo);
			fi.SetIndex(fonts.Size());
			fonts.Add(fi.GetIndex(), fi);

			FontEntityAtom fnt = fi.CreateRecord();

			// Append new child to the end
			AppendChildRecord(fnt);

			// the added font is the last in the list
			return fi;
		}

		public HSLFFontInfo AddFont(InputStream fontData)
		{
			FontHeader fontHeader = new FontHeader();
			InputStream _is = fontHeader.bufferInit(fontData);

			HSLFFontInfo fi = AddFont(fontHeader);

			// always overwrite the font info properties when a font data given
			// as the font info properties are assigned generically when only a typeface is given
			FontEntityAtom fea = fi.GetFontEntityAtom();
			if (fea != null)
			{
				fea.SetCharSet(fontHeader.GetCharsetByte());
				fea.SetPitchAndFamily(FontPitch.getNativeId(fontHeader.GetPitch(), fontHeader.GetFamily()));

				// always activate subsetting
				fea.SetFontFlags(1);
				// true type font and no font substitution
				fea.SetFontType(12);

				Record after = fea;

				int insertIdx = GetFacetIndex(fontHeader.IsItalic(), fontHeader.IsBold());

				FontEmbeddedData newChild = null;
				foreach (FontEmbeddedData fed in fi.GetFacets())
				{

					FontHeader fh = fed.GetFontHeader();
					int curIdx = GetFacetIndex(fh.IsItalic(), fh.IsBold());

					if (curIdx == insertIdx)
					{
						newChild = fed;
						break;
					}
					else if (curIdx > insertIdx)
					{
						// the new facet needs to be inserted before the current facet
						break;
					}

					after = fed;
				}

				if (newChild == null)
				{
					newChild = new FontEmbeddedData();
					AddChildAfter(newChild, after);
					fi.AddFacet(newChild);
				}
				newChild.SetFontData(IOUtils.ToByteArray(_is));
			}
			return fi;
		}

		private static int GetFacetIndex(bool isItalic, bool isBold)
		{
			return (isItalic ? 2 : 0) | (isBold ? 1 : 0);
		}


		/**
		 * Lookup a FontInfo object by its typeface
		 *
		 * @param typeface the full font name
		 *
		 * @return the HSLFFontInfo for the given name or {@code null} if not found
		 */
		public HSLFFontInfo GetFontInfo(string typeface)
		{
			return GetFontInfo(typeface, null);
		}

		/**
		 * Lookup a FontInfo object by its typeface
		 *
		 * @param typeface the full font name
		 * @param charset the charset
		 *
		 * @return the HSLFFontInfo for the given name or {@code null} if not found
		 */
		public HSLFFontInfo GetFontInfo(string typeface, FontCharset charset)
		{
			return fonts.Values.FirstOrDefault(FindFont(typeface, charset));
		}

		private static Func<HSLFFontInfo, bool> FindFont(string typeface, FontCharset charset)
		{
			return (fi) => typeface.Equals(fi.GetTypeface()) && (charset == null || charset.Equals(fi.GetCharset()));
		}

		/**
		 * Lookup a FontInfo object by its internal font index
		 *
		 * @param index the internal font index
		 *
		 * @return the HSLFFontInfo for the given index or {@code null} if not found
		 */
		public HSLFFontInfo GetFontInfo(int index)
		{
			foreach (HSLFFontInfo fi in fonts.Values)
			{
				if (fi.GetIndex() == index)
				{
					return fi;
				}
			}
			return null;
		}

		/**
		 * @return the number of registered fonts
		 */
		public int GetNumberOfFonts()
		{
			return fonts.Size();
		}

		public List<HSLFFontInfo> GetFonts()
		{
			return fonts.Values.ToList();
		}

		//@Override
		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return null;
		}

		public override bool IsAnAtom()
		{
			return true;
		}
	}
}