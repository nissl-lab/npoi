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

namespace NPOI.HWPF.UserModel
{

    using NPOI.HWPF.Model;
    using NPOI.HWPF.UserModel;

    public class ListEntry: Paragraph
    {
        //private static POILogger log = POILogFactory.GetLogger(ListEntry.class);

        ListLevel _level;
        ListFormatOverrideLevel _overrideLevel;

        internal ListEntry(PAPX papx, Range parent, ListTables tables)
            : base(papx, parent)
        {

            if (tables != null)
            {
                ListFormatOverride override1 = tables.GetOverride(_props.GetIlfo());
                _overrideLevel = override1.GetOverrideLevel(_props.GetIlvl());
                _level = tables.GetLevel(override1.GetLsid(), _props.GetIlvl());
            }
            else
            {
                //log.log(POILogger.WARN, "No ListTables found for ListEntry - document probably partly corrupt, and you may experience problems");
            }
        }

        public override int Type
        {
            get
            {
                return Range.TYPE_LISTENTRY;
            }
        }
    }
}

