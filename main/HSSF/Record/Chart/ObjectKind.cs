/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record.Chart
{
    public enum ObjectKind : short {
        AxisGroup = 0,
        AttachedLabelRecord = 0x2,
        Axis = 0x4,
        ChartGroup = 0x5,
        DatRecord = 0x6,
        Frame = 0x7,
        Legend = 0x9,
        LegendException = 0xA,
        Series = 0xC,
        Sheet = 0xD,
        DataFormatRecord = 0xE,
        DropBarRecord = 0xF
    }
}
