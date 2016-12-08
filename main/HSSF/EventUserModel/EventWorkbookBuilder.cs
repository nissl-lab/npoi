/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.HSSF.EventUserModel
{
    using System.Collections;

    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using System.Collections.Generic;

    /// <summary>
    /// When working with the EventUserModel, if you want to
    /// Process formulas, you need an instance of
    /// Workbook to pass to a HSSFWorkbook,
    /// to finally give to HSSFFormulaParser,
    /// and this will build you stub ones.
    /// Since you're working with the EventUserModel, you
    /// wouldn't want to Get a full Workbook and
    ///  HSSFWorkbook, as they would eat too much memory.
    /// Instead, you should collect a few key records as they
    /// go past, then call this once you have them to build a
    /// stub Workbook, and from that a stub
    /// HSSFWorkbook, to use with the HSSFFormulaParser.
    /// The records you should collect are:
    /// ExternSheetRecord
    /// BoundSheetRecord
    /// You should probably also collect SSTRecord,
    /// but it's not required to pass this in.
    /// To help, this class includes a HSSFListener wrapper
    /// that will do the collecting for you.
    /// </summary>
    public class EventWorkbookBuilder
    {
        /// <summary>
        /// Creates a stub Workbook from the supplied records,
        /// suitable for use with the {@link HSSFFormulaParser}
        /// </summary>
        /// <param name="externs">The ExternSheetRecords in your file</param>
        /// <param name="bounds">The BoundSheetRecords in your file</param>
        /// <param name="sst">TThe SSTRecord in your file.</param>
        /// <returns>A stub Workbook suitable for use with HSSFFormulaParser</returns>
        public static InternalWorkbook CreateStubWorkbook(ExternSheetRecord[] externs,
                BoundSheetRecord[] bounds, SSTRecord sst)
        {
            List<Record> wbRecords = new List<Record>();

            // Core Workbook records go first
            if (bounds != null)
            {
                for (int i = 0; i < bounds.Length; i++)
                {
                    wbRecords.Add(bounds[i]);
                }
            }
            if (sst != null)
            {
                wbRecords.Add(sst);
            }

            // Now we can have the ExternSheetRecords,
            //  preceded by a SupBookRecord
            if (externs != null)
            {
                wbRecords.Add(SupBookRecord.CreateInternalReferences(
                        (short)externs.Length));
                for (int i = 0; i < externs.Length; i++)
                {
                    wbRecords.Add(externs[i]);
                }
            }

            // Finally we need an EoF record
            wbRecords.Add(EOFRecord.instance);

            return InternalWorkbook.CreateWorkbook(wbRecords);
        }

        /// <summary>
        /// Creates a stub workbook from the supplied records,
        /// suitable for use with the HSSFFormulaParser
        /// </summary>
        /// <param name="externs">The ExternSheetRecords in your file</param>
        /// <param name="bounds">A stub Workbook suitable for use with HSSFFormulaParser</param>
        /// <returns>A stub Workbook suitable for use with {@link HSSFFormulaParser}</returns>
        public static InternalWorkbook CreateStubWorkbook(ExternSheetRecord[] externs,
                BoundSheetRecord[] bounds)
        {
            return CreateStubWorkbook(externs, bounds, null);
        }


        /// <summary>
        /// A wrapping HSSFListener which will collect
        /// BoundSheetRecords and {@link ExternSheetRecord}s as
        /// they go past, so you can Create a Stub {@link Workbook} from
        /// them once required.
        /// </summary>
        public class SheetRecordCollectingListener : IHSSFListener
        {
            private IHSSFListener childListener;
            private ArrayList boundSheetRecords = new ArrayList();
            private ArrayList externSheetRecords = new ArrayList();
            private SSTRecord sstRecord = null;

            /// <summary>
            /// Initializes a new instance of the <see cref="SheetRecordCollectingListener"/> class.
            /// </summary>
            /// <param name="childListener">The child listener.</param>
            public SheetRecordCollectingListener(IHSSFListener childListener)
            {
                this.childListener = childListener;
            }


            /// <summary>
            /// Gets the bound sheet records.
            /// </summary>
            /// <returns></returns>
            public BoundSheetRecord[] GetBoundSheetRecords()
            {
                return (BoundSheetRecord[])boundSheetRecords.ToArray(
                        typeof(BoundSheetRecord)
                );
            }
            /// <summary>
            /// Gets the extern sheet records.
            /// </summary>
            /// <returns></returns>
            public ExternSheetRecord[] GetExternSheetRecords()
            {
                return (ExternSheetRecord[])externSheetRecords.ToArray(
                        typeof(ExternSheetRecord)
                );
            }
            /// <summary>
            /// Gets the SST record.
            /// </summary>
            /// <returns></returns>
            public SSTRecord GetSSTRecord()
            {
                return sstRecord;
            }

            /// <summary>
            /// Gets the stub HSSF workbook.
            /// </summary>
            /// <returns></returns>
            public HSSFWorkbook GetStubHSSFWorkbook()
            {
	            // Create a base workbook
		            HSSFWorkbook wb = HSSFWorkbook.Create(GetStubWorkbook());
		            // Stub the sheets, so sheet name lookups work
		            foreach (BoundSheetRecord bsr in boundSheetRecords) {
		                wb.CreateSheet(bsr.Sheetname);
		            }
		            // Ready for Formula use!
		            return wb;
            }
            /// <summary>
            /// Gets the stub workbook.
            /// </summary>
            /// <returns></returns>
            public InternalWorkbook GetStubWorkbook()
            {
                return CreateStubWorkbook(
                        GetExternSheetRecords(), GetBoundSheetRecords(),
                        GetSSTRecord()
                );
            }


            /// <summary>
            /// Process this record ourselves, and then
            /// pass it on to our child listener
            /// </summary>
            /// <param name="record">The record.</param>
            public void ProcessRecord(Record record)
            {
                // Handle it ourselves
                ProcessRecordInternally(record);

                // Now pass on to our child
                childListener.ProcessRecord(record);
            }

            /// <summary>
            /// Process the record ourselves, but do not
            /// pass it on to the child Listener.
            /// </summary>
            /// <param name="record">The record.</param>
            public void ProcessRecordInternally(Record record)
            {
                if (record is BoundSheetRecord)
                {
                    boundSheetRecords.Add(record);
                }
                else if (record is ExternSheetRecord)
                {
                    externSheetRecords.Add(record);
                }
                else if (record is SSTRecord)
                {
                    sstRecord = (SSTRecord)record;
                }
            }
        }
    }
}