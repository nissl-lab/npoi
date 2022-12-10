using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI.HSLF.Record
{
	/**
 * Interface to define how a record can indicate it cares about what its
 *  parent is, and how it wants to be told which record is its parent.
 */
	public interface ParentAwareRecord
	{
		RecordContainer GetParentRecord();
		void SetParentRecord(RecordContainer parentRecord);
	}
}