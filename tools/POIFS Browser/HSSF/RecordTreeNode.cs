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

/* ================================================================
 * POIFS Browser 
 * Author: NPOI Team
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * Huseyin Tufekcilerli     2008.11         1.0
 * Tony Qu                  2009.2.18       1.2 alpha
 * 
 * ==============================================================*/

using System;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Aggregates;
using NPOI.DDF;
using NPOI.HSSF.Model;

namespace NPOI.Tools.POIFSBrowser
{
    internal class RecordTreeNode : AbstractRecordTreeNode
    {
        public RecordTreeNode(Record record)
        {
            if (record is ObjRecord)
            {
                ObjRecord or = (ObjRecord)record;
                IList subrecs=or.SubRecords;
                foreach (SubRecord sr in subrecs)
                {
                    this.Nodes.Add(new SubRecordTreeNode(sr));
                }
                this.ImageKey = "Folder";
                this.SelectedImageKey = "Folder";
            }
            else if (record is SSTRecord)
            {
                IEnumerator strings = ((SSTRecord)record).GetStrings();

                this.ImageKey = "Folder";
                this.SelectedImageKey = "Folder";

                while (strings.MoveNext())
                {
                    this.Nodes.Add(new UnicodeStringTreeNode((UnicodeString)strings.Current));
                }
            }
            else if (record is DrawingGroupRecord)
            {
                this.ImageKey = "Folder";
                this.SelectedImageKey = "Folder";

                DrawingGroupRecord dgr = (DrawingGroupRecord)record;

                for (int i = 0; i < dgr.EscherRecords.Count; i++)
                {
                    this.Nodes.Add(new EscherRecordTreeNode((NPOI.DDF.EscherRecord)dgr.EscherRecords[i]));
                }
                
            }
            else
            {
                this.ImageKey = "Binary";
            }
            if (record is UnknownRecord)
            {
                this.Text = UnknownRecord.GetBiffName(record.Sid) + "*";
            }
            else
            {
                this.Text = record.GetType().Name;
            }
            this.Record = record;
            
        }

        public override byte[] GetBytes()
        {
            return ((Record)this.Record).Serialize();
        }

        public override bool HasBinary
        {
            get { return true; }
        }
    }
}
