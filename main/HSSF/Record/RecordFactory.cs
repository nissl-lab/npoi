
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
/* ================================================================
 * About NPOI
 * POI Version: 3.8 beta4
 * Date: 2012-02-15
 * 
 * ==============================================================*/

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
    using System.Globalization;

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
        private static int NUM_RECORDS = 512;
        private static Type[] recordClasses;
        #region inner Record Creater
        private interface I_RecordCreator
        {
            Record Create(RecordInputStream in1);

            Type GetRecordClass();
        }
        private class ReflectionConstructorRecordCreator : I_RecordCreator
        {

            private ConstructorInfo _c;
            public ReflectionConstructorRecordCreator(ConstructorInfo c)
            {
                _c = c;
            }
            public Record Create(RecordInputStream in1)
            {
                Object[] args = { in1 };
                try
                {
                    return (Record)_c.Invoke(args);
                }
                catch (Exception e)
                {
                    throw new RecordFormatException("Unable to construct record instance", e.InnerException);
                }
            }
            public Type GetRecordClass()
            {
                return _c.DeclaringType;
            }
        }
        /**
         * A "create" method is used instead of the usual constructor if the created record might
         * be of a different class to the declaring class.
         */
        private class ReflectionMethodRecordCreator : I_RecordCreator
        {

            private MethodInfo _m;
            public ReflectionMethodRecordCreator(MethodInfo m)
            {
                _m = m;
            }
            public Record Create(RecordInputStream in1)
            {
                Object[] args = { in1 };
                try
                {
                    return (Record)_m.Invoke(null, args);
                }
                catch (Exception e)
                {
                    throw new RecordFormatException("Unable to construct record instance", e.InnerException);
                }
            }
            public Type GetRecordClass()
            {
                return _m.DeclaringType;
            }
        }
        #endregion

        private static Type[] CONSTRUCTOR_ARGS = new Type[] { typeof(RecordInputStream), };


        static RecordFactory()
        {
            recordClasses = new Type[]
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
                typeof(AlRunsRecord),
                //typeof(CodeNameRecord),
		        typeof(CodepageRecord),
		        typeof(ColumnInfoRecord),
		        typeof(ContinueRecord),
		        typeof(CountryRecord),
		        typeof(CRNCountRecord),
		        typeof(CRNRecord),
		        typeof(DateWindow1904Record),
		        typeof(DBCellRecord),
                typeof(DConRefRecord),
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
                typeof(PrintSizeRecord),
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
                typeof(SCLRecord),
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
                typeof(SheetExtRecord),

                // chart records
                //typeof(AlRunsRecord),
                typeof(AreaFormatRecord),
                typeof(AreaRecord),
                typeof(AttachedLabelRecord),
                typeof(AxcExtRecord),
                typeof(AxisLineRecord),
                typeof(AxisParentRecord),
                typeof(AxisRecord),
                typeof(AxesUsedRecord),
                typeof(BarRecord),
                typeof(BeginRecord),
                typeof(BopPopCustomRecord),
                typeof(BopPopRecord),
                typeof(BRAIRecord),
                typeof(CatLabRecord),
                typeof(CatSerRangeRecord),
                typeof(Chart3DBarShapeRecord),
                typeof(Chart3dRecord),
                typeof(ChartEndObjectRecord),
                typeof(ChartFormatRecord),
                typeof(ChartFRTInfoRecord),
                //typeof(ChartRecord),
                typeof(ChartStartObjectRecord),
                typeof(CrtLayout12ARecord),
                typeof(CrtLayout12Record),
                typeof(CrtLineRecord),
                typeof(CrtLinkRecord),
                typeof(CrtMlFrtContinueRecord),
                typeof(CrtMlFrtRecord),
                typeof(DataFormatRecord),
                typeof(DataLabExtContentsRecord),
                typeof(DataLabExtRecord),
                typeof(DatRecord),
                typeof(DefaultTextRecord),
                typeof(DropBarRecord),
                typeof(EndBlockRecord),
                typeof(EndRecord),
                typeof(Fbi2Record),
                typeof(FbiRecord),
                typeof(FontXRecord),
                typeof(FrameRecord),
                typeof(FrtFontListRecord),
                typeof(GelFrameRecord),
                typeof(IFmtRecordRecord),
                typeof(LegendExceptionRecord),
                typeof(LegendRecord),
                typeof(LineFormatRecord),
                typeof(MarkerFormatRecord),
                typeof(ObjectLinkRecord),
                typeof(PicFRecord),
                typeof(PieFormatRecord),
                typeof(PieRecord),
                typeof(PlotAreaRecord),
                typeof(PlotGrowthRecord),
                typeof(PosRecord),
                typeof(RadarAreaRecord),
                typeof(RadarRecord),
                typeof(RichTextStreamRecord),
                typeof(ScatterRecord),
                typeof(SerAuxErrBarRecord),
                typeof(SerAuxTrendRecord),
                typeof(SerFmtRecord),
                typeof(SeriesIndexRecord),
                typeof(SeriesListRecord),
                //typeof(SeriesRecord),
                //typeof(SeriesTextRecord),
                //typeof(SeriesToChartGroupRecord),
                typeof(SerParentRecord),
                typeof(SerToCrtRecord),
                typeof(ShapePropsStreamRecord),
                typeof(ShtPropsRecord),
                typeof(StartBlockRecord),
                typeof(SurfRecord),
                typeof(TextPropsStreamRecord),
                typeof(TextRecord),
                typeof(TickRecord),
                typeof(UnitsRecord),
                //typeof(ValueRangeRecord),
                typeof(YMultRecord),
                		        


                //typeof(BeginRecord),
                //typeof(ChartFRTInfoRecord),
                //typeof(StartBlockRecord),
                //typeof(EndBlockRecord),
                //typeof(ChartStartObjectRecord),
                //typeof(ChartEndObjectRecord),
                //typeof(CatLabRecord),
                //typeof(EndRecord),
                //typeof(PrintSizeRecord),
                
                //typeof(AreaFormatRecord),
                //typeof(AreaRecord),
                //typeof(AxisLineRecord),
                //typeof(AxcExtRecord),
                //typeof(AxisParentRecord),
                //typeof(AxisRecord),
                //typeof(Chart3DBarShapeRecord),
                //typeof(CrtLinkRecord),
                //typeof(AxisUsedRecord),
                //typeof(BarRecord),
                //typeof(CatSerRangeRecord),
                //typeof(DataFormatRecord),
                //typeof(DataLabExtRecord),
                //typeof(DefaultTextRecord),
                //typeof(FrameRecord),
                //typeof(LegendRecord),
                //typeof(FbiRecord),
                //typeof(FontXRecord),
                //typeof(LineFormatRecord),
                //typeof(BRAIRecord),
                //typeof(IFmtRecordRecord),
                //typeof(ObjectLinkRecord),
                //typeof(PlotAreaRecord),
                //typeof(PlotGrowthRecord),
                //typeof(PosRecord),
                //typeof(SCLRecord),
                //typeof(SerToCrtRecord),
                ////typeof(SeriesIndexRecord),
                //typeof(AttachedLabelRecord),
                //typeof(SeriesListRecord),
                ////typeof(SeriesToChartGroupRecord),
                //typeof(ShtPropsRecord),
                //typeof(TextRecord),
                //typeof(TickRecord),
                //typeof(UnitsRecord),

                //typeof(CrtLayout12Record),
                //typeof(Chart3dRecord),
                //typeof(PieRecord),
                //typeof(PieFormatRecord),
                //typeof(CrtLayout12ARecord),
                //typeof(MarkerFormatRecord),
                //typeof(ChartFormatRecord),
                //typeof(SeriesIndexRecord),
                //typeof(GelFrameRecord),
                //typeof(PicFRecord),
                //typeof(CrtMlFrtRecord),
                //typeof(CrtMlFrtContinueRecord),
                //typeof(ShapePropsStreamRecord),
                //end

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

            _recordCreatorsById = RecordsToMap(recordClasses);
        }
        //private static Hashtable recordsMap;
        /**
	     * cache of the recordsToMap();
	     */
        private static Dictionary<short, I_RecordCreator> _recordCreatorsById = null;//RecordsToMap(recordClasses);

        private static short[] _allKnownRecordSIDs;
        /**
         * Debug / diagnosis method<br/>
         * Gets the POI implementation class for a given <tt>sid</tt>.  Only a subset of the any BIFF
         * records are actually interpreted by POI.  A few others are known but not interpreted
         * (see {@link UnknownRecord#getBiffName(int)}).
         * @return the POI implementation class for the specified record <tt>sid</tt>.
         * <code>null</code> if the specified record is not interpreted by POI.
         */
        public static Type GetRecordClass(int sid)
        {
            I_RecordCreator rc = _recordCreatorsById[(short)sid];
            if (rc == null)
            {
                return null;
            }
            return rc.GetRecordClass();
        }
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
        [Obsolete]
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
            if (_recordCreatorsById.ContainsKey(in1.Sid))
            {
                I_RecordCreator constructor = _recordCreatorsById[in1.Sid];
                return constructor.Create(in1);
            }
            else
            {
                return new UnknownRecord(in1);
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
            if (_allKnownRecordSIDs == null)
            {
                short[] results = new short[_recordCreatorsById.Count];
                int i = 0;

                foreach (KeyValuePair<short, I_RecordCreator> kv in _recordCreatorsById)
                {
                    results[i++] = kv.Key;
                }
                Array.Sort(results);
                _allKnownRecordSIDs = results;
            }

            return (short[])_allKnownRecordSIDs.Clone();
        }

        private static Dictionary<short, I_RecordCreator> RecordsToMap(Type[] records)
        {
            Dictionary<short, I_RecordCreator> result = new Dictionary<short, I_RecordCreator>();
            Hashtable uniqueRecClasses = new Hashtable(records.Length * 3 / 2);

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

                short sid = 0;
                try
                {
                    sid = (short)recClass.GetField("sid").GetValue(null);
                }
                catch (Exception ArgumentException)
                {
                    throw new RecordFormatException(
                        "Unable to determine record types", ArgumentException);
                }
                if (result.ContainsKey(sid))
                {
                    Type prevClass = result[sid].GetRecordClass();
                    throw new RuntimeException("duplicate record sid 0x" + sid.ToString("X", CultureInfo.CurrentCulture)
                            + " for classes (" + recClass.Name + ") and (" + prevClass.Name + ")");
                }
                result[sid] = GetRecordCreator(recClass);
            }
            return result;
        }
        [Obsolete]
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

        private static I_RecordCreator GetRecordCreator(Type recClass)
        {
            try
            {
                ConstructorInfo constructor;
                constructor = recClass.GetConstructor(CONSTRUCTOR_ARGS);
                if (constructor != null)
                    return new ReflectionConstructorRecordCreator(constructor);
            }
            catch (Exception)
            {
                // fall through and look for other construction methods
            }
            try
            {
                MethodInfo m = recClass.GetMethod("Create", CONSTRUCTOR_ARGS);
                return new ReflectionMethodRecordCreator(m);
            }
            catch (Exception)
            {
                throw new RuntimeException("Failed to find constructor or create method for (" + recClass.Name + ").");
            }
        }

    }
}
