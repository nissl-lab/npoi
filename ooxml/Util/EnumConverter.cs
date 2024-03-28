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

using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Reflection;

namespace NPOI.Util
{
    public static class EnumConverter
    {       
        public static ST_Jc ValueOf(ParagraphAlignment val)
        {
            switch (val)
            {
                case ParagraphAlignment.BOTH: return ST_Jc.both;
                case ParagraphAlignment.CENTER: return ST_Jc.center;
                case ParagraphAlignment.DISTRIBUTE: return ST_Jc.distribute;
                case ParagraphAlignment.HIGH_KASHIDA: return ST_Jc.highKashida;
                case ParagraphAlignment.LOW_KASHIDA: return ST_Jc.lowKashida;
                case ParagraphAlignment.MEDIUM_KASHIDA: return ST_Jc.mediumKashida;
                case ParagraphAlignment.NUM_TAB: return ST_Jc.numTab;
                case ParagraphAlignment.RIGHT: return ST_Jc.right;
                case ParagraphAlignment.THAI_DISTRIBUTE: return ST_Jc.thaiDistribute;
                default: return ST_Jc.left;
            }
        }
        public static ParagraphAlignment ValueOf(ST_Jc val)
        {
            switch (val)
            {
                case ST_Jc.both: return ParagraphAlignment.BOTH;
                case ST_Jc.center: return ParagraphAlignment.CENTER;
                case ST_Jc.distribute: return ParagraphAlignment.DISTRIBUTE;

                default: return ParagraphAlignment.LEFT;
            }
        }
        public static T ValueOf<T, F>(F val)
        {
            string value = Enum.GetName(val.GetType(), val);
            return (T)Enum.Parse(typeof(T), value, true);
        }
    }
}
