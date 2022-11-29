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

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	/**
 * A special (and dangerous) kind of Record Container, for which other
 *  Atoms care about where this one lives on disk.
 * Will track its position on disk.
 */
	public abstract class PositionDependentRecordContainer: RecordContainer , PositionDependentRecord
	{
		private int sheetId; // Found from PersistPtrHolder

		/**
		 * Fetch our sheet ID, as found from a PersistPtrHolder.
		 * Should match the RefId of our matching SlidePersistAtom
		 */
		public int GetSheetId() { return sheetId; }

		/**
		 * Set our sheet ID, as found from a PersistPtrHolder
		 */
		public void SetSheetId(int id) { sheetId = id; }


		/** Our location on the disk, as of the last write out */
		private int myLastOnDiskOffset;

		/** Fetch our location on the disk, as of the last write out */
		public int GetLastOnDiskOffset() { return myLastOnDiskOffset; }

		/**
		 * Update the Record's idea of where on disk it lives, after a write out.
		 * Use with care...
		 */
		public void SetLastOnDiskOffset(int offset)
		{
			myLastOnDiskOffset = offset;
		}

		/**
		 * Since we're a container, we don't mind if other records move about.
		 * If we're told they have, just return straight off.
		 */
		public void UpdateOtherRecordReferences(Dictionary<int, int> oldToNewReferencesLookup)
		{
		}
	}
}
