
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.Record
{
    using System;
    using System.IO;
    using System.Reflection;
    using System.Collections;
    using System.Collections.Generic;
    using NPOI.HSSF.Record.Chart;
    using NPOI.HSSF.Record.PivotTable;
    using NPOI.HSSF.Record.AutoFilter;
    using NPOI.Util;

    /**
     * Title:  Record Factory
     * Description:  Takes a stream and outputs an array of Record objects.
     *
     * @deprecated use {@link org.apache.poi.hssf.eventmodel.EventRecordFactory} instead
     * @see org.apache.poi.hssf.eventmodel.EventRecordFactory
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Csaba Nagy (ncsaba at yahoo dot com)
     */

    public class RecordFactory
    {
        private static int NUM_RECORDS = 10000;
        private static Type[] records;

        static RecordFactory()
        {
            records = new Type[]
            {
		        typeof(ArrayRecord),
                typeof(AutoFilterInfoRecord),
		        typeof(BackupRecord),
		        typeof(BlankRecord),
		        typeof(BOFRecord),
		        typeof(BookBoolRecord),
		        typeof(BoolErrRecord),
		        typeof(BottomMarginRecord),
		        typeof(BoundSheetRecord),
		        typeof(CalcCountRecord),
		        typeof(CalcModeRecord),
		        typeof(CFHeaderRecord),
		        typeof(CFRuleRecord),
                typeof(ChartRecord),
                typeof(ChartTitleFormatRecord),
                //typeof(CodeNameRecord),
		        typeof(CodepageRecord),

		        typeof(ColumnInfoRecord),
		        typeof(ContinueRecord),
		        typeof(CountryRecord),
		        typeof(CRNCountRecord),
		        typeof(CRNRecord),
		        typeof(DateWindow1904Record),
		        typeof(DBCellRecord),
		        typeof(DefaultColWidthRecord),
		        typeof(DefaultRowHeightRecord),
		        typeof(DeltaRecord),
		        typeof(DimensionsRecord),
		        typeof(DrawingGroupRecord),
		        typeof(DrawingRecord),
		        typeof(DrawingSelectionRecord),
		        typeof(DSFRecord),
		        typeof(DVALRecord),
		        typeof(DVRecord),
		        typeof(EOFRecord),
		        typeof(ExtendedFormatRecord),
		        typeof(ExternalNameRecord),
		        typeof(ExternSheetRecord),
		        typeof(ExtSSTRecord),
		        typeof(FilePassRecord),
		        typeof(FileSharingRecord),
		        typeof(FnGroupCountRecord),
		        typeof(FontRecord),
		        typeof(FooterRecord),
		        typeof(FormatRecord),
		        typeof(FormulaRecord),
		        typeof(GridsetRecord),
		        typeof(GutsRecord),
		        typeof(HCenterRecord),
		        typeof(HeaderRecord),
                typeof(HeaderFooterRecord),
		        typeof(HideObjRecord),
		        typeof(HorizontalPageBreakRecord),
		        typeof(HyperlinkRecord),
		        typeof(IndexRecord),
		        typeof(InterfaceEndRecord),
		        typeof(InterfaceHdrRecord),
		        typeof(IterationRecord),
		        typeof(LabelRecord),
		        typeof(LabelSSTRecord),
		        typeof(LeftMarginRecord),
		        typeof(MergeCellsRecord),
		        typeof(MMSRecord),
		        typeof(MulBlankRecord),
		        typeof(MulRKRecord),
		        typeof(NameRecord),
                typeof(NameCommentRecord),
		        typeof(NoteRecord),
		        typeof(NumberRecord),
		        typeof(ObjectProtectRecord),
		        typeof(ObjRecord),
		        typeof(PaletteRecord),
		        typeof(PaneRecord),
		        typeof(PasswordRecord),
		        typeof(PasswordRev4Record),
		        typeof(PrecisionRecord),
		        typeof(PrintGridlinesRecord),
		        typeof(PrintHeadersRecord),
		        typeof(PrintSetupRecord),
		        typeof(ProtectionRev4Record),
		        typeof(ProtectRecord),
		        typeof(RecalcIdRecord),
		        typeof(RefModeRecord),
		        typeof(RefreshAllRecord),
		        typeof(RightMarginRecord),
		        typeof(RKRecord),
		        typeof(RowRecord),
		        typeof(SaveRecalcRecord),
		        typeof(ScenarioProtectRecord),
		        typeof(SelectionRecord),
                typeof(SeriesRecord),
                typeof(SeriesTextRecord),
		        typeof(SharedFormulaRecord),
		        typeof(SSTRecord),
		        typeof(StringRecord),
		        typeof(StyleRecord),
		        typeof(SupBookRecord),
		        typeof(TabIdRecord),
		        typeof(TableRecord),
                typeof(TableStylesRecord),
		        typeof(TextObjectRecord),
		        typeof(TopMarginRecord),
		        typeof(UncalcedRecord),
		        typeof(UseSelFSRecord),
                typeof(UserSViewBegin),
		        typeof(UserSViewEnd),
                typeof(ValueRangeRecord),
		        typeof(VCenterRecord),
		        typeof(VerticalPageBreakRecord),
		        typeof(WindowOneRecord),
		        typeof(WindowProtectRecord),
		        typeof(WindowTwoRecord),
		        typeof(WriteAccessRecord),
		        typeof(WriteProtectRecord),
		        typeof(WSBoolRecord),

                // chart records
		        typeof(BeginRecord),
		        typeof(ChartFRTInfoRecord),
		        typeof(ChartStartBlockRecord),
		        typeof(ChartEndBlockRecord),
                //TODO typeof(ChartFormatRecord),
		        typeof(ChartStartObjectRecord),
		        typeof(ChartEndObjectRecord),
                typeof(CatLabRecord),
		        typeof(EndRecord),
		        typeof(PrintSizeRecord),
                
                typeof(AreaFormatRecord),
                typeof(AreaRecord),
                typeof(AxisLineFormatRecord),
                typeof(AxisOptionsRecord),
                typeof(AxisParentRecord),
                typeof(AxisRecord),
                typeof(Chart3DBarShapeRecord),
                //typeof(CrtLinkRecord),
                typeof(AxisUsedRecord),
                typeof(BarRecord),
                typeof(CategorySeriesAxisRecord),
                typeof(DataFormatRecord),
                typeof(DataLabelExtensionRecord),
                typeof(DefaultDataLabelTextPropertiesRecord),
                typeof(FrameRecord),
                typeof(LegendRecord),
                typeof(FontBasisRecord),
                typeof(FontIndexRecord),
                typeof(LineFormatRecord),
                typeof(LinkedDataRecord),
                typeof(NumberFormatIndexRecord),
                typeof(ObjectLinkRecord),
                typeof(PlotAreaRecord),
                typeof(PlotGrowthRecord),
                //typeof(PosRecord),
                typeof(SCLRecord),
                typeof(SeriesChartGroupIndexRecord),
                //typeof(SeriesIndexRecord),
                typeof(SeriesLabelsRecord),
                typeof(SeriesListRecord),
		        typeof(SeriesToChartGroupRecord),
                typeof(SheetPropertiesRecord),
                typeof(TextRecord),
                typeof(TickRecord),
                typeof(UnitsRecord),

          		// pivot table records
		        typeof(DataItemRecord),
		        typeof(ExtendedPivotTableViewFieldsRecord),
		        typeof(PageItemRecord),
		        typeof(StreamIDRecord),
		        typeof(ViewDefinitionRecord), 
		        typeof(ViewFieldsRecord),
		        typeof(ViewSourceRecord),

                //autofilter
                typeof(AutoFilterRecord),
                typeof(FilterModeRecord),

                typeof(Excel9FileRecord)
            };

            recordsMap=RecordsToMap(records);
        }
        private static Hashtable recordsMap;

        /**
         * Changes the default capacity (10000) to handle larger files
         */

        public static void SetCapacity(int capacity)
        {
            NUM_RECORDS = capacity;
        }

        /**
         * Create an array of records from an input stream
         *
         * @param in the InputStream from which the records will be
         *           obtained
         *
         * @return an array of Records Created from the InputStream
         *
         * @exception RecordFormatException on error Processing the
         *            InputStream
         */

        public static List<Record> CreateRecords(Stream in1)
        {
            List<Record> records = new List<Record>(NUM_RECORDS);


            RecordFactoryInputStream recStream = new RecordFactoryInputStream(in1, true);

            Record record;
            while ((record = recStream.NextRecord()) != null)
            {
                records.Add(record);
            }

            return records;
        }

        private static void AddAll(List<Record> destList, Record[] srcRecs)
        {
            for (int i = 0; i < srcRecs.Length; i++)
            {
                destList.Add(srcRecs[i]);
            }
        }


        public static Record[] CreateRecord(RecordInputStream in1)
        {
            Record record = CreateSingleRecord(in1);
            if (record is DBCellRecord)
            {
                // Not needed by POI.  Regenerated from scratch by POI when spreadsheet is written
                return new Record[] { null, };
            }
            if (record is RKRecord)
            {
                return new Record[] { ConvertToNumberRecord((RKRecord)record), };
            }
            if (record is MulRKRecord)
            {
                return ConvertRKRecords((MulRKRecord)record);
            }
            return new Record[] { record, };
        }
        /**
 * Converts a {@link MulBlankRecord} into an equivalent array of {@link BlankRecord}s
 */
        public static BlankRecord[] ConvertBlankRecords(MulBlankRecord mbk)
        {
            BlankRecord[] mulRecs = new BlankRecord[mbk.NumColumns];
            for (int k = 0; k < mbk.NumColumns; k++)
            {
                BlankRecord br = new BlankRecord();

                br.Column = k + mbk.FirstColumn;
                br.Row = mbk.Row;
                br.XFIndex = mbk.GetXFAt(k);
                mulRecs[k] = br;
            }
            return mulRecs;
        }

        public static Record CreateSingleRecord(RecordInputStream in1)
        {
            ConstructorInfo constructor = (ConstructorInfo)recordsMap[in1.Sid];

            if (constructor == null)
            {
                return new UnknownRecord(in1);
            }

            try
            {
                return (Record)constructor.Invoke(new object[] { in1 });
            }
            catch (TargetInvocationException e)
            {
                throw new RecordFormatException("Unable to construct record instance: "+in1.Sid, e.InnerException);
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (MethodAccessException)
            {
                throw;
            }
        }
        /// <summary>
        /// RK record is a slightly smaller alternative to NumberRecord
        /// POI likes NumberRecord better
        /// </summary>
        /// <param name="rk">The rk.</param>
        /// <returns></returns>
        public static NumberRecord ConvertToNumberRecord(RKRecord rk)
        {
            NumberRecord num = new NumberRecord();

            num.Column = (rk.Column);
            num.Row = (rk.Row);
            num.XFIndex = (rk.XFIndex);
            num.Value = (rk.RKNumber);
            return num;
        }
        /// <summary>
        /// Converts a MulRKRecord into an equivalent array of NumberRecords
        /// </summary>
        /// <param name="mrk">The MRK.</param>
        /// <returns></returns>
        public static NumberRecord[] ConvertRKRecords(MulRKRecord mrk)
        {

            NumberRecord[] mulRecs = new NumberRecord[mrk.NumColumns];
            for (int k = 0; k < mrk.NumColumns; k++)
            {
                NumberRecord nr = new NumberRecord();

                nr.Column = ((short)(k + mrk.FirstColumn));
                nr.Row = (mrk.Row);
                nr.XFIndex = (mrk.GetXFAt(k));
                nr.Value = (mrk.GetRKNumberAt(k));
                mulRecs[k] = nr;
            }
            return mulRecs;
        }
        public static short[] GetAllKnownRecordSIDs()
        {
            short[] results = new short[recordsMap.Count];
            int i = 0;

            for (IEnumerator iterator = recordsMap.Keys.GetEnumerator();
                    iterator.MoveNext(); )
            {
                short sid = (short)iterator.Current;

                results[i++] = Convert.ToInt16(sid);
            }
            return results;
        }

        private static Hashtable RecordsToMap(Type[] records)
        {
            Hashtable result = new Hashtable();
            Hashtable uniqueRecClasses = new Hashtable(records.Length * 3 / 2);

            ConstructorInfo constructor;

            for (int i = 0; i < records.Length; i++)
            {

                Type recClass = records[i];
                if (!typeof(Record).IsAssignableFrom(recClass))
                {
                    throw new Exception("Invalid record sub-class (" + recClass.Name + ")");
                }
                if (recClass.IsAbstract)
                {
                    throw new Exception("Invalid record class (" + recClass.Name + ") - must not be abstract");
                }
                if (uniqueRecClasses.Contains(recClass))
                {
                    throw new Exception("duplicate record class (" + recClass.Name + ")");
                }
                uniqueRecClasses.Add(recClass, recClass);

                short sid;
                try
                {
                    sid = (short)recClass.GetField("sid").GetValue(null);
                    constructor = recClass.GetConstructor(new Type[]
                    {
                        typeof(RecordInputStream)
                    });
                }
                catch (Exception ArgumentException)
                {
                    throw new RecordFormatException(
                        "Unable to determine record types", ArgumentException);
                }
                result[sid]= constructor;
            }
            return result;
        }
        private static void CheckZeros(Stream in1, int avail)
        {
            int count = 0;
            while (true)
            {
                int b = (int)in1.ReadByte();
                if (b < 0)
                {
                    break;
                }
                if (b != 0)
                {
                    Console.Error.WriteLine(HexDump.ByteToHex(b));
                }
                count++;
            }
            if (avail != count)
            {
                Console.Error.WriteLine("avail!=count (" + avail + "!=" + count + ").");
            }
        }

    }
}
