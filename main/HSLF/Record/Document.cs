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

using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class Document : PositionDependentRecordContainer
	{
		private byte[] _header;
		private static long _type = 1000;

		// Links to our more interesting children
		private DocumentAtom documentAtom;
		private Environment environment;
		private PPDrawingGroup ppDrawing;
		private SlideListWithText[] slwts;
		private ExObjList exObjList; // Can be null

		/**
		 * Returns the DocumentAtom of this Document
		 */
		public DocumentAtom GetDocumentAtom() { return documentAtom; }

		/**
		 * Returns the Environment of this Notes, which lots of
		 * settings for the document in it
		 */
		public Environment GetEnvironment() { return environment; }

		/**
		 * Returns the PPDrawingGroup, which holds an Escher Structure
		 * that contains information on pictures in the slides.
		 */
		public PPDrawingGroup GetPPDrawingGroup() { return ppDrawing; }

		/**
		 * Returns the ExObjList, which holds the references to
		 * external objects used in the slides. This may be null, if
		 * there are no external references.
		 *
		 * @param create if true, create an ExObjList if it doesn't exist
		 */
		public ExObjList GetExObjList(bool create)
		{
			if (exObjList == null && create)
			{
				exObjList = new ExObjList();
				AddChildAfter(exObjList, GetDocumentAtom());
			}
			return exObjList;
		}

		/**
		 * Returns all the SlideListWithTexts that are defined for
		 *  this Document. They hold the text, and some of the text
		 *  properties, which are referred to by the slides.
		 * This will normally return an array of size 2 or 3
		 */
		public SlideListWithText[] GetSlideListWithTexts() { return slwts; }

		/**
		 * Returns the SlideListWithText that deals with the
		 *  Master Slides
		 */
		public SlideListWithText GetMasterSlideListWithText()
		{
			foreach (SlideListWithText slwt in slwts)
			{
				if (slwt.GetInstance() == SlideListWithText.MASTER)
				{
					return slwt;
				}
			}
			return null;
		}

		/**
		 * Returns the SlideListWithText that deals with the
		 *  Slides, or null if there isn't one
		 */
		public SlideListWithText GetSlideSlideListWithText()
		{
			foreach (SlideListWithText slwt in slwts)
			{
				if (slwt.GetInstance() == SlideListWithText.SLIDES)
				{
					return slwt;
				}
			}
			return null;
		}
		/**
		 * Returns the SlideListWithText that deals with the
		 *  notes, or null if there isn't one
		 */
		public SlideListWithText GetNotesSlideListWithText()
		{
			foreach (SlideListWithText slwt in slwts)
			{
				if (slwt.GetInstance() == SlideListWithText.NOTES)
				{
					return slwt;
				}
			}
			return null;
		}


		/**
		 * Set things up, and find our more interesting children
		 */
		/* package */
		Document(byte[] source, int start, int len)
		{
			// Grab the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Find our children
			_children = Record.FindChildRecords(source, start + 8, len - 8);

			// Our first one should be a document atom
			if (!(_children[0] is DocumentAtom))
			{
				throw new InvalidOperationException("The first child of a Document must be a DocumentAtom");
			}
			documentAtom = (DocumentAtom)_children[0];

			// Find how many SlideListWithTexts we have
			// Also, grab the Environment and PPDrawing records
			//  on our way past
			int slwtcount = 0;
			for (int i = 1; i < _children.Length; i++)
			{
				if (_children[i] is SlideListWithText)
				{
					slwtcount++;
				}
				if (_children[i] is Environment)
				{
					environment = (Environment)_children[i];
				}
				if (_children[i] is PPDrawingGroup)
				{
					ppDrawing = (PPDrawingGroup)_children[i];
				}
				if (_children[i] is ExObjList)
				{
					exObjList = (ExObjList)_children[i];
				}
			}

			// You should only every have 1, 2 or 3 SLWTs
			//  (normally it's 2, or 3 if you have notes)
			// Complain if it's not
			if (slwtcount == 0)
			{
				//LOG.atWarn().log("No SlideListWithText's found - there should normally be at least one!");
			}
			if (slwtcount > 3)
			{
				//LOG.atWarn().log("Found {} SlideListWithTexts - normally there should only be three!", box(slwtcount));
			}

			// Now grab all the SLWTs
			slwts = new SlideListWithText[slwtcount];
			slwtcount = 0;
			for (int i = 1; i < _children.Length; i++)
			{
				if (_children[i] is SlideListWithText)
				{
					slwts[slwtcount] = (SlideListWithText)_children[i];
					slwtcount++;
				}
			}
		}

		/**
		 * Adds a new SlideListWithText record, at the appropriate
		 *  point in the child records.
		 */
		public void AddSlideListWithText(SlideListWithText slwt)
		{
			// The new SlideListWithText should go in
			//  just before the EndDocumentRecord
			Record endDoc = _children[_children.Length - 1];
			if (endDoc.GetRecordType() == RecordTypes.RoundTripCustomTableStyles12.typeID)
			{
				// last record can optionally be a RoundTripCustomTableStyles12Atom
				endDoc = _children[_children.Length - 2];
			}
			if (endDoc.GetRecordType() != RecordTypes.EndDocument.typeID)
			{
				throw new InvalidOperationException("The last child record of a Document should be EndDocument, but it was " + endDoc);
			}

			// Add in the record
			AddChildBefore(slwt, endDoc);

			// Updated our cached list of SlideListWithText records
			int newSize = slwts.Length + 1;
			SlideListWithText[] nl = new SlideListWithText[newSize];
			Array.Copy(slwts, 0, nl, 0, slwts.Length);
			nl[nl.Length - 1] = slwt;
			slwts = nl;
		}

		public void RemoveSlideListWithText(SlideListWithText slwt)
		{
			List<SlideListWithText> lst = new List<SlideListWithText>();
			foreach (SlideListWithText s in slwts)
			{
				if (s != slwt) lst.Add(s);
				else
				{
					RemoveChild(slwt);
				}
			}
			slwts = lst.ToArray(new SlideListWithText[0]);
		}

		/**
		 * We are of type 1000
		 */
		public long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		public void WriteOut(OutputStream _out)
		{
			WriteOut(_header[0], _header[1], _type, _children, _out);
		}

		public override int GetLastOnDiskOffset() 
		{
			throw new NotImplementedException();
		}

		public override void SetLastOnDiskOffset(int offset)
		{
			throw new NotImplementedException();
		}

		public override void UpdateOtherRecordReferences(Dictionary<int, int> oldToNewReferencesLookup)
		{
			throw new NotImplementedException();
		}

		public override bool IsAnAtom()
		{
			throw new NotImplementedException();
		}

		public override long GetRecordType()
		{
			throw new NotImplementedException();
		}

		public override Record[] GetChildRecords()
		{
			throw new NotImplementedException();
		}

		public override void WriteOut(BinaryWriter o)
		{
			throw new NotImplementedException();
		}

		public override IDictionary<string, Func<object>> GetGenericProperties()
		{
			throw new NotImplementedException();
		}
	}
}
