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

using NPOI.HSSF.UserModel;
using NPOI.SL.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.WP.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.UserModel
{
	public class HSLFNotes: HSLFSheet, Notes<HSLFShape, HSLFTextParagraph>
	{
		private List<List<HSLFTextParagraph>> _paragraphs = new List<List<HSLFTextParagraph>>();

		/**
		 * Constructs a Notes Sheet from the given Notes record.
		 * Initialises TextRuns, to provide easier access to the text
		 *
		 * @param notes the Notes record to read from
		 */
		public HSLFNotes(Record.Notes notes)
			:base(notes, notes.GetNotesAtom().GetSlideID())
		{
			// Now, build up TextRuns from pairs of TextHeaderAtom and
			// one of TextBytesAtom or TextCharsAtom, found inside
			// EscherTextboxWrapper's in the PPDrawing
			foreach (List<HSLFTextParagraph> l in HSLFTextParagraph.FindTextParagraphs(GetPPDrawing(), this))
			{
				if (!_paragraphs.Contains(l)) _paragraphs.Add(l);
			}

			if (_paragraphs.Count==0)
			{
				//LOG.atWarn().log("No text records found for notes sheet");
			}
		}

		/**
		 * Returns an array of all the TextParagraphs found
		 */
		//@Override
	public override List<List<HSLFTextParagraph>> GetTextParagraphs()
		{
			return _paragraphs;
		}

		/**
		 * Return <code>null</code> - Notes Masters are not yet supported
		 */
		//@Override
	public override HSLFMasterSheet GetMasterSheet()
		{
			return null;
		}

		/**
		 * Header / Footer settings for this slide.
		 *
		 * @return Header / Footer settings for this slide
		 */
		//@Override
	public HeadersFooters GetHeadersFooters()
		{
			return new HeadersFooters(this, HeadersFootersContainer.NotesHeadersFootersContainer);
		}


		//@Override
	public HSLFPlaceholderDetails getPlaceholderDetails(Placeholder placeholder)
		{
			if (placeholder == null)
			{
				return null;
			}

			if (placeholder.nativeEnum == "HEADER" || placeholder.nativeEnum == "FOOTER")
			{
				return new HSLFPlaceholderDetails(this, placeholder);
			}
			else
			{
				return base.GetPlaceholderDetails(placeholder);
			}
		}
	}
}
