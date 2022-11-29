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

using NPOI.SS.Formula.Functions;
using System;

namespace NPOI.HSLF.Record
{
	/**
	 * Abstract class which all atom records will extend.
	 */
	public abstract class RecordAtom: Record
	{
		//arbitrarily selected; may need to increase
		private static int DEFAULT_MAX_RECORD_LENGTH = 1_000_000;
		private static int MAX_RECORD_LENGTH = DEFAULT_MAX_RECORD_LENGTH;

		/**
		 * @param length the max record length allowed for RecordAtom
		 */
		public static void SetMaxRecordLength(int length)
		{
			MAX_RECORD_LENGTH = length;
		}

		/**
		 * @return the max record length allowed for RecordAtom
		 */
		public static int GetMaxRecordLength()
		{
			return MAX_RECORD_LENGTH;
		}

		/**
		 * We are an atom
		 */
		public override bool IsAnAtom() { return true; }

		/**
		 * We're an atom, returns null
		 */
		public override Record[] GetChildRecords() { return null; }
	}
}