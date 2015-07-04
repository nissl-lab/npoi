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

namespace NPOI.HSSF.Record.Aggregates
{
    using System;
    using System.Text;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula;
    using System.Collections.Generic;
    using NPOI.SS.Util;
    using NPOI.SS.Formula.PTG;

    /// <summary>
    /// 
    /// </summary>
    /// CFRecordsAggregate - aggregates Conditional Formatting records CFHeaderRecord
    /// and number of up to three CFRuleRecord records toGether to simplify
    /// access to them.
    /// @author Dmitriy Kumshayev
    public class CFRecordsAggregate : RecordAggregate
    {
        /** Excel allows up to 3 conditional formating rules */
        private const int MAX_97_2003_CONDTIONAL_FORMAT_RULES = 3;

        public const short sid = -2008; // not a real BIFF record

        //private static POILogger log = POILogFactory.GetLogger(typeof(CFRecordsAggregate));

        private CFHeaderRecord header;

        /** List of CFRuleRecord objects */
        private List<CFRuleRecord> rules;

        private CFRecordsAggregate(CFHeaderRecord pHeader, CFRuleRecord[] pRules)
        {
            if (pHeader == null)
            {
                throw new ArgumentException("header must not be null");
            }
            if (pRules == null)
            {
                throw new ArgumentException("rules must not be null");
            }
            if (pRules.Length > MAX_97_2003_CONDTIONAL_FORMAT_RULES)
            {
                Console.WriteLine("Excel versions before 2007 require that "
                    + "No more than " + MAX_97_2003_CONDTIONAL_FORMAT_RULES
                    + " rules may be specified, " + pRules.Length + " were found,"
                    + " this file will cause problems with old Excel versions");
            }
            header = pHeader;
            rules = new List<CFRuleRecord>(3);
            for (int i = 0; i < pRules.Length; i++)
            {
                rules.Add(pRules[i]);
            }
        }


        public CFRecordsAggregate(CellRangeAddress[] regions, CFRuleRecord[] rules)
            : this(new CFHeaderRecord(regions, rules.Length), rules)
        {

        }

        /// <summary>
        /// Create CFRecordsAggregate from a list of CF Records
        /// </summary>
        /// <param name="rs">list of Record objects</param>
        public static CFRecordsAggregate CreateCFAggregate(RecordStream rs)
        {
            Record rec = rs.GetNext();
            if (rec.Sid != CFHeaderRecord.sid)
            {
                throw new InvalidOperationException("next record sid was " + rec.Sid
                        + " instead of " + CFHeaderRecord.sid + " as expected");
            }

            CFHeaderRecord header = (CFHeaderRecord)rec;
            int nRules = header.NumberOfConditionalFormats;

            CFRuleRecord[] rules = new CFRuleRecord[nRules];
            for (int i = 0; i < rules.Length; i++)
            {
                rules[i] = (CFRuleRecord)rs.GetNext();
            }

            return new CFRecordsAggregate(header, rules);
        }
        /// <summary>
        /// Create CFRecordsAggregate from a list of CF Records
        /// </summary>
        /// <param name="recs">list of Record objects</param>
        /// <param name="pOffset">position of CFHeaderRecord object in the list of Record objects</param>
        public static CFRecordsAggregate CreateCFAggregate(IList recs, int pOffset)
        {
            Record rec = (Record)recs[pOffset];
            if (rec.Sid != CFHeaderRecord.sid)
            {
                throw new InvalidOperationException("next record sid was " + rec.Sid
                        + " instead of " + CFHeaderRecord.sid + " as expected");
            }

            CFHeaderRecord header = (CFHeaderRecord)rec;
            int nRules = header.NumberOfConditionalFormats;

            CFRuleRecord[] rules = new CFRuleRecord[nRules];
            int offset = pOffset;
            int countFound = 0;
            while (countFound < rules.Length)
            {
                offset++;
                if (offset >= recs.Count)
                {
                    break;
                }
                rec = (Record)recs[offset];
                if (rec is CFRuleRecord)
                {
                    rules[countFound] = (CFRuleRecord)rec;
                    countFound++;
                }
                else
                {
                    break;
                }
            }

            if (countFound < nRules)
            { // TODO -(MAR-2008) can this ever happen? Write junit 

                //if (log.Check(POILogger.DEBUG))
                //{
                //    log.Log(POILogger.DEBUG, "Expected  " + nRules + " Conditional Formats, "
                //            + "but found " + countFound + " rules");
                //}
                header.NumberOfConditionalFormats = (nRules);
                CFRuleRecord[] lessRules = new CFRuleRecord[countFound];
                Array.Copy(rules, 0, lessRules, 0, countFound);
                rules = lessRules;
            }
            return new CFRecordsAggregate(header, rules);
        }
        public override void VisitContainedRecords(RecordVisitor rv)
        {
            rv.VisitRecord(header);
            for (int i = 0; i < rules.Count; i++)
            {
                CFRuleRecord rule = rules[i];
                rv.VisitRecord(rule);
            }
        }

        /// <summary>
        /// Create a deep Clone of the record
        /// </summary>
        public CFRecordsAggregate CloneCFAggregate()
        {

            CFRuleRecord[] newRecs = new CFRuleRecord[rules.Count];
            for (int i = 0; i < newRecs.Length; i++)
            {
                newRecs[i] = (CFRuleRecord)GetRule(i).Clone();
            }
            return new CFRecordsAggregate((CFHeaderRecord)header.Clone(), newRecs);
        }

        public override short Sid
        {
            get { return sid; }
        }

        /// <summary>
        /// called by the class that is responsible for writing this sucker.
        /// Subclasses should implement this so that their data is passed back in a
        /// byte array.
        /// </summary>
        /// <param name="offset">The offset to begin writing at</param>
        /// <param name="data">The data byte array containing instance data</param>
        /// <returns> number of bytes written</returns>
        public override int Serialize(int offset, byte[] data)
        {
            int nRules = rules.Count;
            header.NumberOfConditionalFormats = (nRules);

            int pos = offset;

            pos += header.Serialize(pos, data);
            for (int i = 0; i < nRules; i++)
            {
                pos += GetRule(i).Serialize(pos, data);
            }
            return pos - offset;
        }

        public CFHeaderRecord Header
        {
            get { return header; }
        }

        private void CheckRuleIndex(int idx)
        {
            if (idx < 0 || idx >= rules.Count)
            {
                throw new ArgumentException("Bad rule record index (" + idx
                        + ") nRules=" + rules.Count);
            }
        }
        public CFRuleRecord GetRule(int idx)
        {
            CheckRuleIndex(idx);
            return rules[idx];
        }
        public void SetRule(int idx, CFRuleRecord r)
        {
            CheckRuleIndex(idx);
            rules[idx] = r;
        }

        /**
         * @return <c>false</c> if this whole {@link CFHeaderRecord} / {@link CFRuleRecord}s should be deleted
         */
        public bool UpdateFormulasAfterCellShift(FormulaShifter shifter, int currentExternSheetIx)
        {
            CellRangeAddress[] cellRanges = header.CellRanges;
            bool changed = false;
            List<CellRangeAddress> temp = new List<CellRangeAddress>();
            for (int i = 0; i < cellRanges.Length; i++)
            {
                CellRangeAddress craOld = cellRanges[i];
                CellRangeAddress craNew = ShiftRange(shifter, craOld, currentExternSheetIx);
                if (craNew == null)
                {
                    changed = true;
                    continue;
                }
                temp.Add(craNew);
                if (craNew != craOld)
                {
                    changed = true;
                }
            }

            if (changed)
            {
                int nRanges = temp.Count;
                if (nRanges == 0)
                {
                    return false;
                }
                CellRangeAddress[] newRanges = new CellRangeAddress[nRanges];
                newRanges = temp.ToArray();
                header.CellRanges = (newRanges);
            }

            for (int i = 0; i < rules.Count; i++)
            {
                CFRuleRecord rule = rules[i];
                Ptg[] ptgs;
                ptgs = rule.ParsedExpression1;
                if (ptgs != null && shifter.AdjustFormula(ptgs, currentExternSheetIx))
                {
                    rule.ParsedExpression1 = (ptgs);
                }
                ptgs = rule.ParsedExpression2;
                if (ptgs != null && shifter.AdjustFormula(ptgs, currentExternSheetIx))
                {
                    rule.ParsedExpression2 = (ptgs);
                }
            }
            return true;
        }
        private static CellRangeAddress ShiftRange(FormulaShifter shifter, CellRangeAddress cra, int currentExternSheetIx)
        {
            // FormulaShifter works well in terms of Ptgs - so convert CellRangeAddress to AreaPtg (and back) here
            AreaPtg aptg = new AreaPtg(cra.FirstRow, cra.LastRow, cra.FirstColumn, cra.LastColumn, false, false, false, false);
            Ptg[] ptgs = { aptg, };

            if (!shifter.AdjustFormula(ptgs, currentExternSheetIx))
            {
                return cra;
            }
            Ptg ptg0 = ptgs[0];
            if (ptg0 is AreaPtg)
            {
                AreaPtg bptg = (AreaPtg)ptg0;
                return new CellRangeAddress(bptg.FirstRow, bptg.LastRow, bptg.FirstColumn, bptg.LastColumn);
            }
            if (ptg0 is AreaErrPtg)
            {
                return null;
            }
            throw new InvalidCastException("Unexpected shifted ptg class (" + ptg0.GetType().Name + ")");
        }
        public void AddRule(CFRuleRecord r)
        {
            if (rules.Count >= MAX_97_2003_CONDTIONAL_FORMAT_RULES)
            {
                Console.WriteLine("Excel versions before 2007 cannot cope with"
                    + " any more than " + MAX_97_2003_CONDTIONAL_FORMAT_RULES
                    + " - this file will cause problems with old Excel versions");
            }
            rules.Add(r);
            header.NumberOfConditionalFormats = (rules.Count);
        }
        public int NumberOfRules
        {
            get { return rules.Count; }
        }

        /**
         *  @return sum of sizes of all aggregated records
         */
        //public override int RecordSize
        //{
        //    get
        //    {
        //        int size = 0;
        //        if (header != null)
        //        {
        //            size += header.RecordSize;
        //        }
        //        if (rules != null)
        //        {
        //            for (IEnumerator irecs = rules.GetEnumerator(); irecs.MoveNext(); )
        //            {
        //                size += ((Record)irecs.Current).RecordSize;
        //            }
        //        }
        //        return size;
        //    }
        //}

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CF]\n");
            if (header != null)
            {
                buffer.Append(header.ToString());
            }
            for (int i = 0; i < rules.Count; i++)
            {
                CFRuleRecord cfRule = rules[i];
                if (cfRule != null)
                {
                    buffer.Append(cfRule.ToString());
                }
            }
            buffer.Append("[/CF]\n");
            return buffer.ToString();
        }
    }
}