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

namespace NPOI.HSLF.Record
{
    using System;
    using System.Collections;

    /**
     * A special (and dangerous) kind of Record Container, for which other
     *  Atoms care about where this one lives on disk.
     * Will track its position on disk.
     *
     * @author Nick Burch
     */

    public abstract class PositionDependentRecordContainer : RecordContainer, PositionDependentRecord
    {
        private int sheetId; // Found from PersistPtrHolder

        /**
         * Fetch our sheet ID, as found from a PersistPtrHolder.
         * Should match the RefId of our matching SlidePersistAtom
         */
        public int SheetId
        {
            get { return sheetId; }
            set { sheetId = value; }
        }


        /** Our location on the disk, as of the last write out */
        protected int myLastOnDiskOffset;

        /** Fetch our location on the disk, as of the last write out */
        public int LastOnDiskOffset
        {
            get { return myLastOnDiskOffset; }
            set { myLastOnDiskOffset = value; }
        }

        /**
         * Since we're a Container, we don't mind if other records move about.
         * If we're told they have, just return straight off.
         */
        public void UpdateOtherRecordReferences(Hashtable oldToNewReferencesLookup)
        {
            return;
        }
    }
}