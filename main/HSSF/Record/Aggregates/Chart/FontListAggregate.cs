using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    /// <summary>
    /// FONTLIST = FrtFontList StartObject *(Font [Fbi]) EndObject
    /// </summary>
    public class FontListAggregate : RecordAggregate
    {
        private FrtFontListRecord frtFontList = null;
        private ChartStartObjectRecord startObject = null;
        private Dictionary<FontRecord, FbiRecord> dicFonts = new Dictionary<FontRecord, FbiRecord>();
        private ChartEndObjectRecord endObject = null;
        public FontListAggregate(RecordStream rs)
        {
            frtFontList = (FrtFontListRecord)rs.GetNext();
            startObject = (ChartStartObjectRecord)rs.GetNext();
            FontRecord f = null;
            FbiRecord fbi = null;
            while (rs.PeekNextSid() == FontRecord.sid)
            {
                f = (FontRecord)rs.GetNext();
                if (rs.PeekNextSid() == FbiRecord.sid)
                {
                    fbi = (FbiRecord)rs.GetNext();
                }
                else
                {
                    fbi = null;
                }
                dicFonts.Add(f, fbi);
            }

            endObject = (ChartEndObjectRecord)rs.GetNext();
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(frtFontList);
            rv.VisitRecord(startObject);

            foreach (KeyValuePair<FontRecord, FbiRecord> kv in dicFonts)
            {
                rv.VisitRecord(kv.Key);
                if (kv.Value != null)
                    rv.VisitRecord(kv.Value);
            }

            rv.VisitRecord(endObject);
        }
    }
}
