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
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	/**
 * Environment, which contains lots of settings for the document.
 */
	public class Environment : PositionDependentRecordContainer
	{
		private byte[] _header;
		private static long _type = 1010;

		// Links to our more interesting children
		private FontCollection fontCollection;
		//master style for text with type=TextHeaderAtom.OTHER_TYPE
		private TxMasterStyleAtom txmaster;

		/**
		 * Returns the FontCollection of this Environment
		 */
		public FontCollection GetFontCollection() { return fontCollection; }


		/**
		 * Set things up, and find our more interesting children
		 */
		protected Environment(byte[] source, int start, int len)
		{
			// Grab the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Find our children
			_children = Record.FindChildRecords(source, start + 8, len - 8);

			// Find our FontCollection record
			foreach (Record child in _children)
			{
				if (child is FontCollection)
				{
					fontCollection = (FontCollection)child;
				}
				else if (child is TxMasterStyleAtom)
				{
					txmaster = (TxMasterStyleAtom)child;
				}
			}

			if (fontCollection == null)
			{
				throw new InvalidOperationException("Environment didn't contain a FontCollection record!");
			}
		}

		public TxMasterStyleAtom GetTxMasterStyleAtom()
		{
			return txmaster;
		}

		/**
		 * We are of type 1010
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
