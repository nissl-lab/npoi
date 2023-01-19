/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/

using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class InteractiveInfo: RecordContainer
	{
		private byte[] _header;
		private static long _type = RecordTypes.InteractiveInfo.typeID;

		// Links to our more interesting children
		private InteractiveInfoAtom infoAtom;

		/**
		 * Returns the InteractiveInfoAtom of this InteractiveInfo
		 */
		public InteractiveInfoAtom GetInteractiveInfoAtom() { return infoAtom; }

		/**
		 * Set things up, and find our more interesting children
		 */
		protected InteractiveInfo(byte[] source, int start, int len)
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
			// First child should be the InteractiveInfoAtom
			if (_children == null || _children.Length == 0 || !(_children[0] is InteractiveInfoAtom)) {
				//LOG.atWarn().log("First child record wasn't a InteractiveInfoAtom - leaving this atom in an invalid state...");
				return;
			}

			infoAtom = (InteractiveInfoAtom)_children[0];
		}

		/**
		 * Create a new InteractiveInfo, with blank fields
		 */
		public InteractiveInfo()
		{
			_header = new byte[8];
			_children = new Record[1];

			// Setup our header block
			_header[0] = 0x0f; // We are a container record
			LittleEndian.PutShort(_header, 2, (short)_type);

			// Setup our child records
			infoAtom = new InteractiveInfoAtom();
			_children[0] = infoAtom;
		}

		/**
		 * We are of type 4802
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
