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

using NPOI.Util;
namespace NPOI.HWPF.Model
{

    public class FieldDescriptor
    {
        byte _fieldBoundaryType;
        byte _info;
        private static BitField fZombieEmbed = BitFieldFactory.GetInstance(0x02);
        private static BitField fResultDiry = BitFieldFactory.GetInstance(0x04);
        private static BitField fResultEdited = BitFieldFactory.GetInstance(0x08);
        private static BitField fLocked = BitFieldFactory.GetInstance(0x10);
        private static BitField fPrivateResult = BitFieldFactory.GetInstance(0x20);
        private static BitField fNested = BitFieldFactory.GetInstance(0x40);
        private static BitField fHasSep = BitFieldFactory.GetInstance(0x80);


        public FieldDescriptor()
        {
        }
    }
}

