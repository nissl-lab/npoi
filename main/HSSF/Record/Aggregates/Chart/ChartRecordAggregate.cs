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

using System.Collections.Generic;
using NPOI.HSSF.Record.Chart;

namespace NPOI.HSSF.Record.Aggregates.Chart
{
    public abstract class ChartRecordAggregate : RecordAggregate
    {
        public const string RuleName_DAT = "DAT";
        public const string RuleName_TEXTPROPS = "TEXTPROPS";
        public const string RuleName_SS = "SS";
        public const string RuleName_SHAPEPROPS = "SHAPEPROPS";
        public const string RuleName_SERIESFORMAT = "SERIESFORMAT";
        public const string RuleName_SERIESAXIS = "SERIESAXIS";
        public const string RuleName_LD = "LD";
        public const string RuleName_IVAXIS = "IVAXIS";
        public const string RuleName_GELFRAME = "GELFRAME";
        public const string RuleName_FRAME = "FRAME";
        public const string RuleName_FONTLIST = "FONTLIST";
        public const string RuleName_DVAXIS = "DVAXIS";
        public const string RuleName_DROPBAR = "DROPBAR";
        public const string RuleName_DFTTEXT = "DFTTEXT";
        public const string RuleName_CRTMLFRT = "CRTMLFRT";
        public const string RuleName_CRT = "CRT";
        public const string RuleName_CHARTFOMATS = "CHARTFOMATS";
        public const string RuleName_AXS = "AXS";
        public const string RuleName_AXM = "AXM";
        public const string RuleName_AXISPARENT = "AXISPARENT";
        public const string RuleName_AXES = "AXES";
        public const string RuleName_ATTACHEDLABEL = "ATTACHEDLABEL";
        public const string RuleName_LEGENDEXCEPTION = "LEGENDEXCEPTION";
        public const string RuleName_CHARTSHEET = "CHARTSHEET";

        private string _rulename;

        protected string RuleName
        {
            get { return _rulename; }
            private set { _rulename = value; }
        }
        ChartRecordAggregate _container;
        protected ChartRecordAggregate Container
        {
            get { return _container; }
            private set { _container = value; }
        }
        protected ChartRecordAggregate(string ruleName, ChartRecordAggregate container)
        {
            this.RuleName = ruleName;
            this.Container = container;
        }
        private static StartBlockStack blocks = new StartBlockStack();

        private bool _isInStartObject;
        protected bool IsInStartObject
        {
            get { return _isInStartObject; }
            set { _isInStartObject=value;}
        }
        public const short ChartSpecificFutureRecordLowerSid = 0x800;
        public const short ChartSpecificFutureRecordHigherSid = 0x8FF;
        protected virtual bool ShoudWriteStartBlock()
        {
            return false;
        }
        private bool IsInRule(string ruleName)
        {
            ChartRecordAggregate cra = this;
            while (cra != null)
            {
                if (cra.RuleName == ruleName)
                    return true;
                cra = cra.Container;
            }
            return false;
        }
        protected T GetContainer<T>(string ruleName) where T : ChartRecordAggregate
        {
            ChartRecordAggregate cra = this;
            while (cra != null)
            {
                if (cra.RuleName == ruleName)
                    break;
                cra = cra.Container;
            }
            return cra as T;
        }
        protected void WriteStartBlock(RecordVisitor rv)
        {
            //1
            //A StartBlock record MUST not be written if the record is preceded by a StartObject record 
            //but not preceded by the matching EndObject record. That is, StartBlock and EndBlock pairs 
            //MUST not belong to any collection defined by StartObject and EndObject.
            if (IsInStartObject)
                return;

            StartBlockRecord sbr = null;
            //2
            //If there does not exist a StartBlock record with iObjectKind equal to 0x000D without 
            //a matching EndBlock record, then a corresponding StartBlock record with iObjectKind 
            //equal to 0x000D MUST be written.
            if (blocks.Count == 0)
            {
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Sheet);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }
            //3
            //If the chart-specific future record exists in the sequence of records that conforms to 
            //the DAT rule, and there does not exist a StartBlock record with iObjectKind equal to 0x0006 
            //without a matching EndBlock record, then a corresponding StartBlock record with iObjectKind 
            //equal to 0x0006 MUST be written. If a StartBlock record is written because of rule number 2,
            //then this StartBlock record MUST be written immediately after that record.

            if (IsInRule(RuleName_DAT) && !blocks.Contains(ObjectKind.DatRecord))
            {
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.DatRecord);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }
            //4
            //If the chart-specific future record is in a series, and there does not exist a StartBlock record 
            //with iObjectKind equal to 0x000C without a matching EndBlock record, then a corresponding StartBlock 
            //record with iObjectKind equal to 0x000C and iObjectInstance1 equal to the number of series prior to 
            //this series in the current Sheet MUST be written. If any StartBlock records are written because of 
            //rule number 2 or 3, then this StartBlock record MUST be written immediately after those records.
            if (IsInRule(RuleName_SERIESFORMAT) && !blocks.Contains(ObjectKind.Series))
            {
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Series, 0,
                    GetContainer<SeriesFormatAggregate>(RuleName_SERIESFORMAT).SeriesIndex);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }
            //5
            //If the chart-specific future record exists in the sequence of records that conforms to the SS rule, 
            //and there does not exist a StartBlock record with iObjectKind equal to 0x000E without a matching EndBlock 
            //record, then a corresponding StartBlock record with iObjectKind equal to 0x000E, iObjectContext equal to 
            //the yi field of the DataFormat record in the current SS rule, and iObjectInstance1 equal to the xi field 
            //of the DataFormat record in the current SS rule MUST be written. If any StartBlock records are written 
            //because of rule number 2, 3, or 4, then this StartBlock record MUST be written immediately after those records.
            if (IsInRule(RuleName_SS) && !blocks.Contains(ObjectKind.DataFormatRecord))
            {
                SSAggregate ss = GetContainer<SSAggregate>(RuleName_SS);
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.DataFormatRecord,
                    ss.DataFormat.SeriesIndex, ss.DataFormat.PointNumber);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }
            //6
            //If the chart-specific future record is in a series, and is part of a collection defined by a Begin and End 
            //pair written immediately after a LegendException record, and there does not exist a StartBlock record with 
            //iObjectKind equal to 0x000A without a matching EndBlock record, then a corresponding StartBlock record with 
            //iObjectKind equal to 0x000A and iObjectInstance1 equal to the iss field of the LegendException record in 
            //the series MUST be written. If any StartBlock records are written because of rule number 2, 3, 4, or 5, 
            //then this StartBlock record MUST be written immediately after those records.

            if (IsInRule(RuleName_LEGENDEXCEPTION) && !blocks.Contains(ObjectKind.LegendException))
            {
                SeriesFormatAggregate.LegendExceptionAggregate le =
                    GetContainer<SeriesFormatAggregate.LegendExceptionAggregate>(RuleName_LEGENDEXCEPTION);
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.LegendException, 0,
                    le.LegendException.LegendEntry);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }

            //7
            //If the chart-specific future record is in an axis group. and there does not exist a StartBlock record with 
            //iObjectKind equal to 0x0000 without a matching EndBlock record, then a corresponding StartBlock record with 
            //iObjectKind equal to 0x0000 and iObjectInstance1 equal to the iax field of the AxisParent record of the axis 
            //group MUST be written. If any StartBlock records are written because of rule number 2, 3, 4, 5, or 6, then 
            //this StartBlock record MUST be written immediately after those records.
            if (IsInRule(RuleName_AXISPARENT) && !blocks.Contains(ObjectKind.AxisGroup))
            {
                AxisParentAggregate ap =  GetContainer<AxisParentAggregate>(RuleName_AXISPARENT);
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AxisGroup, 0,
                    ap.AxisParent.AxisType);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }


            //8
            //If the chart-specific future record is in a Chart Group, and there does not exist a StartBlock record with 
            //iObjectKind equal to 0x0005 without a matching EndBlock record, then a corresponding StartBlock record with 
            //iObjectKind equal to 0x0005 and iObjectInstance1 equal to the iax field of the AxisParent record of the axis 
            //group MUST be written. If any StartBlock records are written because of rule number 2, 3, 4, 5, 6, or 7, then 
            //this StartBlock record MUST be written immediately after those records.
            if (IsInRule(RuleName_CRT) && !blocks.Contains(ObjectKind.ChartGroup))
            {
                AxisParentAggregate ap =  GetContainer<AxisParentAggregate>(RuleName_AXISPARENT);
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.ChartGroup, 0,
                    ap.AxisParent.AxisType);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }


            //9
            //If the chart-specific future record is in an axis, and there does not exist a StartBlock record with iObjectKind 
            //equal to 0x0004 without a matching EndBlock record, then:
            if (IsInRule(RuleName_AXES) && !blocks.Contains(ObjectKind.Axis))
            {
                //If the chart-specific future record exists in the sequence of records that conforms to the IVAXIS rule, 
                //then a corresponding StartBlock record with iObjectKind equal to 0x0004 and iObjectInstance1 equal to 
                //0x0000 MUST be written. If any StartBlock records are written because of rule number 2, 3, 4, 5, 6, 7, 
                //or 8, then this StartBlock record MUST be written immediately after those records.
                if (IsInRule(RuleName_IVAXIS))
                {
                    sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Axis, 0, 0);
                    blocks.Push(sbr);
                    rv.VisitRecord(sbr);
                }

                //If the chart-specific future record exists in the sequence of records that conforms to the SERIESAXIS rule, 
                //then a corresponding StartBlock record with iObjectKind equal to 0x0004 and iObjectInstance1 equal to 0x0002 
                //MUST be written. If any StartBlock records are written because of rule number 2, 3, 4, 5, 6, 7, or 8, then 
                //this StartBlock record MUST be written immediately after those records.
                if (IsInRule(RuleName_SERIESAXIS))
                {
                    sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Axis, 0, 2);
                    blocks.Push(sbr);
                    rv.VisitRecord(sbr);
                }

                //If the chart-specific future record exists in the sequence of records that conforms to the DVAXIS rule, and 
                //wType of the Axis record in the sequence of records that conforms to the DVAXIS rule is equal to 0, then a 
                //corresponding StartBlock record with iObjectKind equal to 0x0004 and iObjectInstance1 equal to 0x0001 MUST 
                //be written. If any StartBlock records are written because of rule number 2, 3, 4, 5, 6, 7, or 8, then this 
                //StartBlock record MUST be written immediately after those records.
                if (IsInRule(RuleName_DVAXIS))
                {
                    DVAxisAggregate dva = GetContainer<DVAxisAggregate>(RuleName_DVAXIS);
                    if (dva.Axis.AxisType == AxisRecord.AXIS_TYPE_CATEGORY_OR_X_AXIS)
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Axis, 0, 1);
                    else
                    {
                        //If the chart-specific future record exists in the sequence of records that conforms to the DVAXIS rule, and 
                        //wType of the Axis record in the sequence of records that conforms to the DVAXIS rule is equal to 1, then a 
                        //corresponding StartBlock record with iObjectKind equal to 0x0004 and iObjectInstance1 equal to 0x0003 MUST 
                        //be written. If any StartBlock records are written because of rule number 2, 3, 4, 5, 6, 7, or 8, then this 
                        //StartBlock record MUST be written immediately after those records.
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Axis, 0, 3);
                    }
                    blocks.Push(sbr);
                    rv.VisitRecord(sbr);


                }
            }
            //10
            //If the chart-specific future record exists in the sequence of records that conforms to the DROPBAR rule, and 
            //there does not exist a StartBlock record with iObjectKind equal to 0x000F without a matching EndBlock record, 
            //then a corresponding StartBlock record with iObjectKind equal to 0x000F and iObjectInstance1 equal to one less 
            //than the number of DropBar records written prior to the chart-specific future record in the current Chart Group
            //MUST be written. If any StartBlock records are written because of rule number 2, 3, 4, 5, 6, 7, 8, or 9, then 
            //this StartBlock record MUST be written immediately after those records.
            if (IsInRule(RuleName_DROPBAR) && !blocks.Contains(ObjectKind.DropBarRecord))
            {
                sbr = StartBlockRecord.CreateStartBlock(ObjectKind.DropBarRecord);
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }


            //11
            //If the chart-specific future record is in a legend and there does not exist a StartBlock record with iObjectKind 
            //equal to 0x0009 without a matching EndBlock record, then:
            if (IsInRule(RuleName_LD) && !blocks.Contains(ObjectKind.Legend))
            {
                if (IsInRule(RuleName_CRT))
                {
                    sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Legend, 1);
                    //If the chart-specific future record is in a chart group, then a corresponding StartBlock record with iObjectKind 
                    //equal to 0x0009 and iObjectContext equal to 0x0001 MUST be written. If any StartBlock records are written because
                    //of rule number 2, 3, 4, 5, 6, 7, 8, 9, or 10, then this StartBlock record MUST be written immediately after those 
                    //records.
                }
                else
                {
                    sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Legend, 0);
                    //If the chart-specific future record is not in a chart group, then a corresponding StartBlock record with iObjectKind
                    //equal to 0x0009 and iObjectContext equal to 0x0000 MUST be written. If any StartBlock records are written because 
                    //of rule number 2, 3, 4, 5, 6, 7, 8, 9, or 10, then this StartBlock record MUST be written immediately after those 
                    //records.
                }
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }

            //12
            //If the chart-specific future record is in an attached label, and there does not exist a StartBlock record with iObjectKind 
            //equal to 0x0002 without a matching EndBlock record, then:
            if (IsInRule(RuleName_ATTACHEDLABEL) && !blocks.Contains(ObjectKind.AttachedLabelRecord))
            {
                if (IsInRule(RuleName_DFTTEXT))
                {
                    //If the chart-specific future record exists in the sequence of records that conforms to the DFTTEXT rule of a
                    //chart group, and the id field of the DefaultText record in the sequence of records that conforms to the DFTTEXT 
                    //rule is greater than or equal to 0x0002, then a corresponding StartBlock record with iObjectKind equal to 0x0002,
                    //iObjectContext equal to 0x0002, and iObjectInstance1 equal to 0xFFFF MUST be written. If any StartBlock records 
                    //are written because of rule number 2, 3, 4, 5, 6, 7, 8, 9, 10, or 11, then this StartBlock record MUST be written 
                    //immediately after those records. Else,
                    DFTTextAggregate dft = GetContainer<DFTTextAggregate>(RuleName_DFTTEXT);
                    if (IsInRule(RuleName_CRT) && (int)dft.DefaultText.FormatType >= 2)
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AttachedLabelRecord, 2,
                            unchecked((short)0xFFFF));
                    //If the chart-specific future record exists in the sequence of records that conforms to the DFTTEXT rule of a 
                    //chart group, then a corresponding StartBlock record with iObjectKind equal to 0x0002, iObjectContext equal to
                    //0x0002, and iObjectInstance1 equal to the id field of the DefaultText record in the sequence of records that 
                    //conforms to the DFTTEXT rule MUST be written. If any StartBlock records are written because of rule number 
                    //2, 3, 4, 5, 6, 7, 8, 9, 10, or 11, then this StartBlock record MUST be written immediately after those records. Else,
                    else
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AttachedLabelRecord, 2,
                            (short)dft.DefaultText.FormatType);
                }
                else
                {
                    AttachedLabelAggregate ala = GetContainer<AttachedLabelAggregate>(RuleName_ATTACHEDLABEL);
                    //If the wLinkVar1 of the ObjectLink record of the attached label is equal to 0x0003, then a corresponding 
                    //StartBlock record with iObjectKind equal to 0x0002, iObjectContext equal to 0x0004 and iObjectInstance1 
                    //equal to 0x0000 MUST be written. If any StartBlock records are written because of rules number 2, 3, 4, 
                    //5, 6, 7, 8, 9, 10 or 11, then this StartBlock record MUST be written immediately after those records. Else,
                    if (ala.ObjectLink.Link1 == 3)
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AttachedLabelRecord, 4, 0);
                    //If the wLinkVar1 of the ObjectLink record of the attached label is equal to 0x0002, then a corresponding 
                    //StartBlock record with iObjectKind equal to 0x0002, iObjectContext equal to 0x0004 and iObjectInstance1 
                    //equal to 0x0001 MUST be written. If any StartBlock records are written because of rules number 2, 3, 4, 
                    //5, 6, 7, 8, 9, 10 or 11, then this StartBlock record MUST be written immediately after those records. Else,
                    else if (ala.ObjectLink.Link1 == 2)
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AttachedLabelRecord, 4, 1);
                    //If the wLinkVar1 of the ObjectLink record of the attached label is equal to 0x0007, then a corresponding 
                    //StartBlock record with iObjectKind equal to 0x0002, iObjectContext equal to 0x0004, and iObjectInstance1 
                    //equal to 0x0002 MUST be written. If any StartBlock records are written because of rule number 2, 3, 4, 5,
                    //6, 7, 8, 9, 10, or 11, then this StartBlock record MUST be written immediately after those records. Else,
                    else if (ala.ObjectLink.Link1 == 7)
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AttachedLabelRecord, 4, 2);

                    else if (ala.IsFirst)
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AttachedLabelRecord, 0);
                    //If the chart-specific future record is in the first attached label of a chart sheet, then a corresponding 
                    //StartBlock record with iObjectKind equal to 0x0002 and iObjectContext equal to 0x0000 MUST be written. If
                    //any StartBlock records are written because of rule number 2, 3, 4, 5, 6, 7, 8, 9, 10, or 11, then this 
                    //StartBlock record MUST be written immediately after those records. Else,
                    else
                        sbr = StartBlockRecord.CreateStartBlock(ObjectKind.AttachedLabelRecord, 5, 
                            ala.ObjectLink.Link1, ala.ObjectLink.Link2);
                    //If the chart-specific future record is not in the first attached label of a chart sheet, then a corresponding 
                    //StartBlock record with iObjectKind equal to 0x0002 and iObjectContext equal to 0x0005, iObjectInstance1 
                    //equal to wLinkVar1 of the ObjectLink record of the attached label and iObjectInstance2 equal to wLinkVar2 
                    //of the ObjectLink record of the attached label MUST be written. If any StartBlock records are written 
                    //because of rule number 2, 3, 4, 5, 6, 7, 8, 9, 10, or 11, then this StartBlock record MUST be written 
                    //immediately after those records.
                }
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }

            //13
            //If the chart-specific future record exists in the sequence of records that conforms to the FRAME rule, and there 
            //does not exist a StartBlock record with iObjectKind equal to 0x0007 without a matching EndBlock record, then:
            if (IsInRule(RuleName_FRAME) && !blocks.Contains(ObjectKind.Frame))
            {
                if (IsInRule(RuleName_ATTACHEDLABEL) || IsInRule(RuleName_LD))
                {
                    //If the chart-specific future record is in an attached label or legend, then a corresponding StartBlock record 
                    //with iObjectKind equal to 0x0007, iObjectContext equal to 0x0000, and iObjectInstance1 equal to 0x0000 MUST be 
                    //written. If any StartBlock records are written because of rules number 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, or 12, 
                    //then this StartBlock record MUST be written immediately after those records. Else,
                    sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Frame, 0, 0);
                }
                else if (IsInRule(RuleName_AXES))
                {
                    //If the chart-specific future record exists in the sequence of records that conforms to the AXES rule, then a 
                    //corresponding StartBlock record with iObjectKind equal to 0x0007, iObjectContext equal to 0x0001, and 
                    //iObjectInstance1 equal to 0x0000 MUST be written. If any StartBlock records are written because of rule number 
                    //2, 3, 4, 5, 6, 7, 8, 9, 10, 11, or 12, then this StartBlock record MUST be written immediately after 
                    //those records. Else,
                    sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Frame, 1, 0);
                }
                else if (IsInRule(RuleName_CHARTSHEET))
                {
                    //If the chart-specific future record is in a Sheet, then a corresponding StartBlock record with iObjectKind 
                    //equal to 0x0007, iObjectContext equal to 0x0002, and iObjectInstance1 equal to 0x0000 MUST be written. If any 
                    //StartBlock records are written because of rule number 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, or 12, then this StartBlock
                    //record MUST be written immediately after those records.
                    sbr = StartBlockRecord.CreateStartBlock(ObjectKind.Frame, 2, 0);
                }
                blocks.Push(sbr);
                rv.VisitRecord(sbr);
            }
        }
        protected void WriteEndBlock(RecordVisitor rv)
        {
            if (IsInStartObject)
                return;

            StartBlockRecord sbr = blocks.Peek();
            //If there exists a StartBlock record with iObjectKind equal to 0x0000 without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current Axis Group.
            if (this.RuleName == RuleName_AXISPARENT && sbr.ObjectKind == ObjectKind.AxisGroup)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.AxisGroup));
                blocks.Pop();
                return;
            }

            //If there exists a StartBlock record with iObjectKind equal to 0x0002 without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current AttachedLabel.
            if (this.RuleName == RuleName_ATTACHEDLABEL && sbr.ObjectKind == ObjectKind.AttachedLabelRecord)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.AttachedLabelRecord));
                blocks.Pop();
                return;
            }

            //If there exists a StartBlock record with iObjectKind equal to 0x0004 without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current Axis.
            if ((this.RuleName == RuleName_IVAXIS || this.RuleName==RuleName_DVAXIS||this.RuleName==RuleName_SERIESAXIS)
                && sbr.ObjectKind == ObjectKind.Axis)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.Axis));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x0005 without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current chart group.
            if (this.RuleName == RuleName_CRT && sbr.ObjectKind == ObjectKind.ChartGroup)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.ChartGroup));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x0006 without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the sequence of records 
            //containing the StartBlock and conforming to the DAT rule.
            if (this.RuleName == RuleName_DAT && sbr.ObjectKind == ObjectKind.DatRecord)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.DatRecord));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x0007 without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the sequence of records 
            //containing the StartBlock and conforming to the FRAME rule.
            if (this.RuleName == RuleName_FRAME && sbr.ObjectKind == ObjectKind.Frame)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.Frame));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x0009 without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current Legend.
            if (this.RuleName == RuleName_LD && sbr.ObjectKind == ObjectKind.Legend)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.Legend));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x000A without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current Begin 
            //and End collection that exists immediately after LegendException in the sequence of records conforming 
            //to the SERIESFORMAT rule.
            if (this.RuleName == RuleName_LEGENDEXCEPTION && sbr.ObjectKind == ObjectKind.LegendException)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.LegendException));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x000C without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current Series.
            if (this.RuleName == RuleName_SERIESFORMAT && sbr.ObjectKind == ObjectKind.Series)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.Series));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x000D without a matching EndBlock,
            //then a matching EndBlock record MUST exist immediately before the End record of the current Sheet.
            if (this.RuleName == RuleName_CHARTFOMATS && sbr.ObjectKind == ObjectKind.Sheet)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.Sheet));
                blocks.Pop();
                return;
            }

            //If there exists a StartBlock record with iObjectKind equal to 0x000E without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the current SS production.
            if (this.RuleName == RuleName_SS && sbr.ObjectKind == ObjectKind.DataFormatRecord)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.DataFormatRecord));
                blocks.Pop();
                return;
            }
            //If there exists a StartBlock record with iObjectKind equal to 0x000F without a matching EndBlock, 
            //then a matching EndBlock record MUST exist immediately before the End record of the sequence of 
            //records containing the StartBlock and conforming to the DROPBAR rule.
            if (this.RuleName == RuleName_DROPBAR && sbr.ObjectKind == ObjectKind.DropBarRecord)
            {
                rv.VisitRecord(EndBlockRecord.CreateEndBlock(ObjectKind.DropBarRecord));
                blocks.Pop();
                return;
            }
        }

        private class StartBlockStack
        {
            private List<StartBlockRecord> blockList = new List<StartBlockRecord>(16);
            public void Push(StartBlockRecord item)
            {
                blockList.Add(item);
            }
            public StartBlockRecord Pop()
            {
                if(blockList.Count==0)
                    return null;
                StartBlockRecord item = blockList[blockList.Count - 1];
                blockList.RemoveAt(blockList.Count - 1);
                return item;
            }
            public StartBlockRecord Peek()
            {
                if (blockList.Count == 0)
                    return null;
                return blockList[blockList.Count - 1];
            }
            public bool Contains(ObjectKind objectKind)
            {
                foreach (StartBlockRecord item in blockList)
                {
                    if (item.ObjectKind == objectKind)
                        return true;
                }
                return false;
            }
            public int Count
            {
                get { return blockList.Count; }
            }
        }
    }
}
