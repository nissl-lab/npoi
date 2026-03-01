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


using System;

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.OpenXmlFormats.Dml.Chart;
    using NPOI.Util.Optional;
    using NPOI.XDDF.UserModel.Text;

    public class XDDFLegendEntry : ITextContainer
    {
        private CT_LegendEntry entry;
        internal XDDFLegendEntry(CT_LegendEntry entry)
        {
            this.entry = entry;
        }
        internal CT_LegendEntry GetXmlObject()
        {
            return entry;
        }

        public XDDFTextBody TextBody
        {
            get
            {
                if(entry.IsSetTxPr())
                {
                    return new XDDFTextBody(this, entry.txPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    if(entry.IsSetTxPr())
                    {
                        entry.UnsetTxPr();
                    }
                }
                else
                {
                    entry.txPr = value.GetXmlObject();
                }
            }
        }

        public bool? Delete
        {
            get
            {
                if(entry.IsSetDelete())
                {
                    return entry.delete.val == 1;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                if(!value.HasValue)
                {
                    if(entry.IsSetDelete())
                    {
                        entry.UnsetDelete();
                    }
                }
                else
                {
                    if(entry.IsSetDelete())
                    {
                        entry.delete.val = value.Value ? 1 : 0;
                    }
                    else
                    {
                        entry.AddNewDelete().val = value.Value ? 1 : 0;
                    }
                }
            }
        }

        public long Index
        {
            get
            {
                return entry.idx.val;
            }
            set
            {
                entry.idx.val = (uint) value;
            }
        }


        public void SetExtensionList(XDDFChartExtensionList list)
        {

        }

        public XDDFChartExtensionList ExtensionList
        {
            get
            {
                if(entry.IsSetExtLst())
                {
                    return new XDDFChartExtensionList(entry.extLst);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    if(entry.IsSetExtLst())
                    {
                        entry.UnsetExtLst();
                    }
                }
                else
                {
                    entry.extLst = value.GetXmlObject();
                }
            }
        }

        public Option<R> FindDefinedParagraphProperty<R>(
                Func<CT_TextParagraphProperties, Boolean> isSet,
                Func<CT_TextParagraphProperties, R> Getter) where R : class
        {
            return Option<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public Option<R> FindDefinedRunProperty<R>(
                Func<CT_TextCharacterProperties, Boolean> isSet,
                Func<CT_TextCharacterProperties, R> Getter) where R : class
        {
            return Option<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public ValueOption<R> FindDefinedParagraphValueProperty<R>(
                Func<CT_TextParagraphProperties, Boolean> isSet,
                Func<CT_TextParagraphProperties, R> Getter) where R : struct
        {
            return ValueOption<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public ValueOption<R> FindDefinedRunValueProperty<R>(
                Func<CT_TextCharacterProperties, Boolean> isSet,
                Func<CT_TextCharacterProperties, R> Getter) where R : struct
        {
            return ValueOption<R>.None(); // legend entry has no (indirect) paragraph properties
        }
    }
}


