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

namespace NPOI.HSSF.Model
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using System.Collections.Generic;
    using NPOI.HSSF.Record.PivotTable;

    /**
     * Finds correct insert positions for records in workbook streams<p/>
     * 
     * See OOO excelfileformat.pdf sec. 4.2.5 'Record Order in a BIFF8 Workbook Stream'
     * 
     * @author Josh Micich
     */
    public class RecordOrderer
    {

        // TODO - simplify logic using a generalised record ordering

        private RecordOrderer()
        {
            // no instances of this class
        }
        /**
         * Adds the specified new record in the correct place in sheet records list
         * 
         */
        public static void AddNewSheetRecord(List<RecordBase> sheetRecords, RecordBase newRecord)
        {
            int index = FindSheetInsertPos(sheetRecords, newRecord.GetType());
            sheetRecords.Insert(index, newRecord);
        }

        private static int FindSheetInsertPos(List<RecordBase> records, Type recClass)
        {
            if (recClass == typeof(DataValidityTable))
            {
                return FindDataValidationTableInsertPos(records);
            }
            if (recClass == typeof(MergedCellsTable))
            {
                return FindInsertPosForNewMergedRecordTable(records);
            }
            if (recClass == typeof(ConditionalFormattingTable))
            {
                return FindInsertPosForNewCondFormatTable(records);
            }
            if (recClass == typeof(GutsRecord))
            {
                return GetGutsRecordInsertPos(records);
            }
            if (recClass == typeof(PageSettingsBlock))
            {
                return GetPageBreakRecordInsertPos(records);
            }
            if (recClass == typeof(WorksheetProtectionBlock))
            {
                return GetWorksheetProtectionBlockInsertPos(records);
            }
            throw new InvalidOperationException("Unexpected record class (" + recClass.Name + ")");
        }
        /// <summary>
        /// Finds the index where the protection block should be inserted
        /// </summary>
        /// <param name="records">the records for this sheet</param>
        /// <returns></returns>
        /// <remark>
        /// + BOF
        /// o INDEX
        /// o Calculation Settings Block
        /// o PRINTHEADERS
        /// o PRINTGRIDLINES
        /// o GRIDSET
        /// o GUTS
        /// o DEFAULTROWHEIGHT
        /// o SHEETPR
        /// o Page Settings Block
        /// o Worksheet Protection Block
        /// o DEFCOLWIDTH
        /// oo COLINFO
        /// o SORT
        /// + DIMENSION
        /// </remark>
        private static int GetWorksheetProtectionBlockInsertPos(List<RecordBase> records)
        {
            int i = GetDimensionsIndex(records);
            while (i > 0)
            {
                i--;
                Object rb = records[i];
                if (!IsProtectionSubsequentRecord(rb))
                {
                    return i + 1;
                }
            }
            throw new InvalidOperationException("did not find insert pos for protection block");
        }
        /// <summary>
        /// These records may occur between the 'Worksheet Protection Block' and DIMENSION:
        /// </summary>
        /// <param name="rb"></param>
        /// <returns></returns>
        /// <remarks>
        /// o DEFCOLWIDTH
        /// oo COLINFO
        /// o SORT
        /// </remarks>
        private static bool IsProtectionSubsequentRecord(Object rb)
        {
            if (rb is ColumnInfoRecordsAggregate)
            {
                return true; // oo COLINFO
            }
            if (rb is Record)
            {
                Record record = (Record)rb;
                switch (record.Sid)
                {
                    case DefaultColWidthRecord.sid:
                    case UnknownRecord.SORT_0090:
                        return true;
                }
            }
            return false;
        }
        private static int GetPageBreakRecordInsertPos(List<RecordBase> records)
        {
            int dimensionsIndex = GetDimensionsIndex(records);
            int i = dimensionsIndex - 1;
            while (i > 0)
            {
                i--;
                RecordBase rb = records[i];
                if (IsPageBreakPriorRecord(rb))
                {
                    return i + 1;
                }
            }
            throw new InvalidOperationException("Did not Find insert point for GUTS");
        }
        private static bool IsPageBreakPriorRecord(RecordBase rb)
        {
            if (rb is Record)
            {
                Record record = (Record)rb;
                switch (record.Sid)
                {
                    case BOFRecord.sid:
                    case IndexRecord.sid:
                    // calc settings block
                    case UncalcedRecord.sid:
                    case CalcCountRecord.sid:
                    case CalcModeRecord.sid:
                    case PrecisionRecord.sid:
                    case RefModeRecord.sid:
                    case DeltaRecord.sid:
                    case IterationRecord.sid:
                    case DateWindow1904Record.sid:
                    case SaveRecalcRecord.sid:
                    // end calc settings
                    case PrintHeadersRecord.sid:
                    case PrintGridlinesRecord.sid:
                    case GridsetRecord.sid:
                    case DefaultRowHeightRecord.sid:
                    case UnknownRecord.SHEETPR_0081:
                        return true;
                    // next is the 'Worksheet Protection Block'
                }
            }
            return false;
        }
        /// <summary>
        /// Find correct position to add new CFHeader record
        /// </summary>
        /// <param name="records"></param>
        /// <returns></returns>
        private static int FindInsertPosForNewCondFormatTable(List<RecordBase> records)
        {

            for (int i = records.Count - 2; i >= 0; i--)
            { // -2 to skip EOF record
                Object rb = records[i];
                if (rb is MergedCellsTable)
                {
                    return i + 1;
                }
                if (rb is DataValidityTable)
                {
                    continue;
                }
                Record rec = (Record)rb;
                switch (rec.Sid)
                {
                    case WindowTwoRecord.sid:
                    case SCLRecord.sid:
                    case PaneRecord.sid:
                    case SelectionRecord.sid:
                    case UnknownRecord.STANDARDWIDTH_0099:
                    // MergedCellsTable usually here 
                    case UnknownRecord.LABELRANGES_015F:
                    case UnknownRecord.PHONETICPR_00EF:
                        // ConditionalFormattingTable goes here
                        return i + 1;
                    // HyperlinkTable (not aggregated by POI yet)
                    // DataValidityTable
                }
            }
            throw new InvalidOperationException("Did not Find Window2 record");
        }

        private static int FindInsertPosForNewMergedRecordTable(List<RecordBase> records)
        {
            for (int i = records.Count - 2; i >= 0; i--)
            { // -2 to skip EOF record
                Object rb = records[i];
                if (!(rb is Record))
                {
                    // DataValidityTable, ConditionalFormattingTable, 
                    // even PageSettingsBlock (which doesn't normally appear after 'View Settings')
                    continue;
                }
                Record rec = (Record)rb;
                switch (rec.Sid)
                {
                    // 'View Settings' (4 records) 
                    case WindowTwoRecord.sid:
                    case SCLRecord.sid:
                    case PaneRecord.sid:
                    case SelectionRecord.sid:

                    case UnknownRecord.STANDARDWIDTH_0099:
                        return i + 1;
                }
            }
            throw new InvalidOperationException("Did not Find Window2 record");
        }


        /**
         * Finds the index where the sheet validations header record should be inserted
         * @param records the records for this sheet
         * 
         * + WINDOW2
         * o SCL
         * o PANE
         * oo SELECTION
         * o STANDARDWIDTH
         * oo MERGEDCELLS
         * o LABELRANGES
         * o PHONETICPR
         * o Conditional Formatting Table
         * o Hyperlink Table
         * o Data Validity Table
         * o SHEETLAYOUT
         * o SHEETPROTECTION
         * o RANGEPROTECTION
         * + EOF
         */
        private static int FindDataValidationTableInsertPos(List<RecordBase> records)
        {
            int i = records.Count - 1;
            if (!(records[i] is EOFRecord))
            {
                throw new InvalidOperationException("Last sheet record should be EOFRecord");
            }
            while (i > 0)
            {
                i--;
                RecordBase rb = records[i];
                if (IsDVTPriorRecord(rb))
                {
                    Record nextRec = (Record)records[i + 1];
                    if (!IsDVTSubsequentRecord(nextRec.Sid))
                    {
                        throw new InvalidOperationException("Unexpected (" + nextRec.GetType().Name
                                + ") found after (" + rb.GetType().Name + ")");
                    }
                    return i + 1;
                }
                Record rec = (Record)rb;
                if (!IsDVTSubsequentRecord(rec.Sid))
                {
                    throw new InvalidOperationException("Unexpected (" + rec.GetType().Name
                            + ") while looking for DV Table insert pos");
                }
            }
            return 0;
        }


        private static bool IsDVTPriorRecord(RecordBase rb)
        {
            if (rb is MergedCellsTable || rb is ConditionalFormattingTable)
            {
                return true;
            }
            short sid = ((Record)rb).Sid;
            switch (sid)
            {
                case WindowTwoRecord.sid:
                case SCLRecord.sid:
                case PaneRecord.sid:
                case SelectionRecord.sid:
                case UnknownRecord.STANDARDWIDTH_0099:
                // MergedCellsTable
                case UnknownRecord.LABELRANGES_015F:
                case UnknownRecord.PHONETICPR_00EF:
                // ConditionalFormattingTable
                case HyperlinkRecord.sid:
                case UnknownRecord.QUICKTIP_0800:
                    return true;
            }
            return false;
        }

        private static bool IsDVTSubsequentRecord(short sid)
        {
            switch (sid)
            {
                //case UnknownRecord.SHEETEXT_0862:
                case SheetExtRecord.sid:
                case UnknownRecord.SHEETPROTECTION_0867:
                case UnknownRecord.PLV_MAC:
                //case UnknownRecord.RANGEPROTECTION_0868:
                case FeatRecord.sid: 
                case EOFRecord.sid:
                    return true;
            }
            return false;
        }
        /**
         * DIMENSIONS record is always present
         */
        private static int GetDimensionsIndex(List<RecordBase> records)
        {
            int nRecs = records.Count;
            for (int i = 0; i < nRecs; i++)
            {
                if (records[i] is DimensionsRecord)
                {
                    return i;
                }
            }
            // worksheet stream is seriously broken
            throw new InvalidOperationException("DimensionsRecord not found");
        }

        private static int GetGutsRecordInsertPos(List<RecordBase> records)
        {
            int dimensionsIndex = GetDimensionsIndex(records);
            int i = dimensionsIndex - 1;
            while (i > 0)
            {
                i--;
                RecordBase rb = records[i];
                if (IsGutsPriorRecord(rb))
                {
                    return i + 1;
                }
            }
            throw new InvalidOperationException("Did not Find insert point for GUTS");
        }

        private static bool IsGutsPriorRecord(RecordBase rb)
        {
            if (rb is Record)
            {
                Record record = (Record)rb;
                switch (record.Sid)
                {
                    case BOFRecord.sid:
                    case IndexRecord.sid:
                    // calc settings block
                        case UncalcedRecord.sid:
                        case CalcCountRecord.sid:
                        case CalcModeRecord.sid:
                        case PrecisionRecord.sid:
                        case RefModeRecord.sid:
                        case DeltaRecord.sid:
                        case IterationRecord.sid:
                        case DateWindow1904Record.sid:
                        case SaveRecalcRecord.sid:
                    // end calc settings
                    case PrintHeadersRecord.sid:
                    case PrintGridlinesRecord.sid:
                    case GridsetRecord.sid:
                        return true;
                    // DefaultRowHeightRecord.sid is next
                }
            }
            return false;
        }
        /// <summary>
        /// if the specified record ID terminates a sequence of Row block records
        /// It is assumed that at least one row or cell value record has been found prior to the current 
        /// record
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        public static bool IsEndOfRowBlock(int sid)
        {
            switch (sid)
            {
                case ViewDefinitionRecord.sid:                // should have been prefixed with DrawingRecord (0x00EC), but bug 46280 seems to allow this
                case DrawingRecord.sid:
                case DrawingSelectionRecord.sid:
                case ObjRecord.sid:
                case TextObjectRecord.sid:
                case ColumnInfoRecord.sid: // See Bugzilla 53984
                case GutsRecord.sid:   // see Bugzilla 50426
                case WindowOneRecord.sid:
                // should really be part of workbook stream, but some apps seem to put this before WINDOW2
                case WindowTwoRecord.sid:
                    return true;

                case DVALRecord.sid:
                    return true;
                case EOFRecord.sid:
                    // WINDOW2 should always be present, so shouldn't have got this far
                    throw new InvalidOperationException("Found EOFRecord before WindowTwoRecord was encountered");
            }
            return PageSettingsBlock.IsComponentRecord(sid);
        }
        /// <summary>
        /// Whether the specified record id normally appears in the row blocks section of the sheet records
        /// </summary>
        /// <param name="sid"></param>
        /// <returns></returns>
        public static bool IsRowBlockRecord(int sid)
        {
            switch (sid)
            {
                case RowRecord.sid:

                case BlankRecord.sid:
                case BoolErrRecord.sid:
                case FormulaRecord.sid:
                case LabelRecord.sid:
                case LabelSSTRecord.sid:
                case NumberRecord.sid:
                case RKRecord.sid:

                case ArrayRecord.sid:
                case SharedFormulaRecord.sid:
                case TableRecord.sid:
                    return true;

            }
            return false;
        }
    }
}