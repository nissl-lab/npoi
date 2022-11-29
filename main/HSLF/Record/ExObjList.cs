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

using NPOI.POIFS.Properties;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI.HSLF.Record
{
	/**
 * This class holds the links to exernal objects referenced from the document.
 */
	public class ExObjList : RecordContainer
	{
		private byte[] _header;
		private static long _type = RecordTypes.ExObjList.typeID;

		// Links to our more interesting children
		private ExObjListAtom exObjListAtom;

		/**
		 * Returns the ExObjListAtom of this list
		 */
		public ExObjListAtom GetExObjListAtom() { return exObjListAtom; }

		/**
		 * Returns all the ExHyperlinks
		 */
		public ExHyperlink[] GetExHyperlinks()
		{
			List<ExHyperlink> links = new List<ExHyperlink>();
			foreach (Record child in _children)
			{
				if (child is ExHyperlink)
				{
					links.Add((ExHyperlink)child);
				}
			}

			return links.ToArray();
		}

		/**
		 * Set things up, and find our more interesting children
		 */
		protected ExObjList(byte[] source, int start, int len)
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
			// First child should be the atom
			if (_children[0] is ExObjListAtom)
			{
				exObjListAtom = (ExObjListAtom)_children[0];
			}
			else
			{
				throw new InvalidOperationException("First child record wasn't a ExObjListAtom, was of type " + _children[0].GetRecordType());
			}
		}

		/**
		 * Create a new ExObjList, with blank fields
		 */
		public ExObjList()
		{
			_header = new byte[8];
			_children = new Record[1];

			// Setup our header block
			_header[0] = 0x0f; // We are a container record
			LittleEndian.PutShort(_header, 2, (short)_type);

			// Setup our child records
			_children[0] = new ExObjListAtom();
			FindInterestingChildren();
		}

		/**
		 * We are of type 1033
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

		/**
		 * Lookup a hyperlink by its unique id
		 *
		 * @param id hyperlink id
		 * @return found <code>ExHyperlink</code> or <code>null</code>
		 */
		public ExHyperlink Get(int id)
		{
			foreach (Record child in _children)
			{
				if (child is ExHyperlink)
				{
					ExHyperlink rec = (ExHyperlink)child;
					if (rec.GetExHyperlinkAtom().GetNumber() == id)
					{
						return rec;
					}
				}
			}
			return null;
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