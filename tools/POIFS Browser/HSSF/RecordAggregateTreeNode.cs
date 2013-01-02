using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.Record.Aggregates;
using System.Windows.Forms;
using NPOI.HSSF.Record;
using System.Collections;
using NPOI.SS.Util;

namespace NPOI.Tools.POIFSBrowser
{
    class RecordAggregateTreeNode:AbstractRecordTreeNode
    {
        public RecordAggregateTreeNode(RecordAggregate record)
        {
            this.Record = record;
            this.Text = record.GetType().Name;
            this.ImageKey = "Folder";
            this.SelectedImageKey = "Folder";

            GetChildren();
        }
        public override bool HasBinary
        {
            get { return false; }
        }

        public override byte[] GetBytes()
        {
            throw new NotImplementedException();
        }
        private void GetChildren()
        {
            RecordAggregate record = (RecordAggregate)this.Record;

            if (record is RowRecordsAggregate)
            {

                IEnumerator recordenum = ((RowRecordsAggregate)record).GetEnumerator();
                while (recordenum.MoveNext())
                {
                    if (recordenum.Current is RowRecord)
                    {
                        this.Nodes.Add(new RecordTreeNode((RowRecord)recordenum.Current));
                    }
                }
                CellValueRecordInterface[] valrecs = ((RowRecordsAggregate)record).GetValueRecords();
                for (int j = 0; j < valrecs.Length; j++)
                {
                    CellValueRecordTreeNode cvrtn = new CellValueRecordTreeNode(valrecs[j]);
                    if (valrecs[j] is FormulaRecordAggregate)
                    {
                        FormulaRecordAggregate fra = ((FormulaRecordAggregate)valrecs[j]);
                        cvrtn.ImageKey = "Folder";
                        if (fra.FormulaRecord != null)
                            cvrtn.Nodes.Add(new RecordTreeNode(fra.FormulaRecord));
                        if (fra.StringRecord != null)
                            cvrtn.Nodes.Add(new RecordTreeNode(fra.StringRecord));
                    }
                    this.Nodes.Add(cvrtn);
                }
            }
            else if (record is ColumnInfoRecordsAggregate)
            {
                IEnumerator recordenum = ((ColumnInfoRecordsAggregate)record).GetEnumerator();
                while (recordenum.MoveNext())
                {
                    if (recordenum.Current is ColumnInfoRecord)
                    {
                        this.Nodes.Add(new RecordTreeNode((ColumnInfoRecord)recordenum.Current));
                    }
                }
            }
            else if (record is PageSettingsBlock)
            {
                PageSettingsBlock psb = (PageSettingsBlock)record;
                MockRecordVisitor rv = new MockRecordVisitor();
                psb.VisitContainedRecords(rv);
                foreach (Record rec in rv.Records)
                {
                    this.Nodes.Add(new RecordTreeNode(rec));
                }
            }
            else if (record is MergedCellsTable)
            {
                foreach (CellRangeAddress subRecord in ((MergedCellsTable)record).MergedRegions)
                {
                    this.Nodes.Add(new CellRangeAddressTreeNode(subRecord));
                }
            }
            else if (record is ConditionalFormattingTable)
            {
                ConditionalFormattingTable cft = (ConditionalFormattingTable)record;
                for (int j = 0; j < cft.Count; j++)
                {
                    CFRecordsAggregate cfra = cft.Get(j);

                    AbstractRecordTreeNode headernode = new RecordTreeNode(cfra.Header);
                    this.Nodes.Add(headernode);
                    for (int k = 0; k < cfra.NumberOfRules; k++)
                    {
                        this.Nodes.Add(new RecordTreeNode(cfra.GetRule(k)));
                    }
                }
            }
            else if (record is WorksheetProtectionBlock)
            {
                WorksheetProtectionBlock wpb = (WorksheetProtectionBlock)record;
                MockRecordVisitor rv=new MockRecordVisitor();
                wpb.VisitContainedRecords(rv);
                foreach (Record rec in rv.Records)
                {
                    this.Nodes.Add(new RecordTreeNode(rec));
                }
            }
        }
    }
}
