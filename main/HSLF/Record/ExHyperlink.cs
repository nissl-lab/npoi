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

namespace NPOI.HSLF.Record
{
	/**
 * This class represents the data of a link in the document.
 */
	public class ExHyperlink : RecordContainer
	{
		private static long _type = RecordTypes.ExHyperlink.typeID;

		private byte[] _header;

		// Links to our more interesting children
		private ExHyperlinkAtom linkAtom;
		private CString linkDetailsA;
		private CString linkDetailsB;

		/**
		 * Returns the ExHyperlinkAtom of this link
		 */
		public ExHyperlinkAtom GetExHyperlinkAtom() { return linkAtom; }

		/**
		 * Returns the URL of the link.
		 *
		 * @return the URL of the link
		 */
		public string GetLinkURL()
		{
			return linkDetailsB == null ? null : linkDetailsB.GetText();
		}

		/**
		 * Returns the hyperlink's user-readable name
		 *
		 * @return the hyperlink's user-readable name
		 */
		public string GetLinkTitle()
		{
			return linkDetailsA == null ? null : linkDetailsA.GetText();
		}

		/**
		 * Sets the URL of the link
		 * TODO: Figure out if we should always set both
		 */
		public void SetLinkURL(string url)
		{
			if (linkDetailsB != null)
			{
				linkDetailsB.SetText(url);
			}
		}

		public void SetLinkOptions(int options)
		{
			if (linkDetailsB != null)
			{
				linkDetailsB.SetOptions(options);
			}
		}

		public void SetLinkTitle(string title)
		{
			if (linkDetailsA != null)
			{
				linkDetailsA.SetText(title);
			}
		}

		/**
		 * Get the link details (field A)
		 */
		public string _GetDetailsA()
		{
			return linkDetailsA == null ? null : linkDetailsA.GetText();
		}
		/**
		 * Get the link details (field B)
		 */
		public string _GetDetailsB()
		{
			return linkDetailsB == null ? null : linkDetailsB.GetText();
		}

		/**
		 * Set things up, and find our more interesting children
		 */
		protected ExHyperlink(byte[] source, int start, int len)
		{
			// Grab the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

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

			// First child should be the ExHyperlinkAtom
			Record child = _children[0];
			if (child is ExHyperlinkAtom)
			{
				linkAtom = (ExHyperlinkAtom)child;
			}
			else
			{
				//LOG.atError().log("First child record wasn't a ExHyperlinkAtom, was of type {}", box(child.getRecordType()));
			}

			for (int i = 1; i < _children.Length; i++)
			{
				child = _children[i];
				if (child is CString)
				{
					if (linkDetailsA == null) linkDetailsA = (CString)child;
					else linkDetailsB = (CString)child;
				}
				else
				{
					//LOG.atError().log("Record after ExHyperlinkAtom wasn't a CString, was of type {}", box(child.getRecordType()));
				}

			}
		}

		/**
		 * Create a new ExHyperlink, with blank fields
		 */
		public ExHyperlink()
		{
			_header = new byte[8];
			_children = new Record[3];

			// Setup our header block
			_header[0] = 0x0f; // We are a container record
			LittleEndian.PutShort(_header, 2, (short)_type);

			// Setup our child records
			CString csa = new CString();
			CString csb = new CString();
			csa.SetOptions(0x00);
			csb.SetOptions(0x10);
			_children[0] = new ExHyperlinkAtom();
			_children[1] = csa;
			_children[2] = csb;
			FindInterestingChildren();
		}

		/**
		 * We are of type 4055
		 */
		public override long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		public override void WriteOut(OutputStream _out)
		{
			WriteOut(_header[0], _header[1], _type, _children, _out);
		}

		public override bool IsAnAtom()
		{
			return true;
		}

		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			throw new NotImplementedException();
		}
	}
}