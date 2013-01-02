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
     * A special (and dangerous) kind of Record Atom that cares about where
     *  it lives on the disk, or who has other Atoms that care about where
     *  this is on the disk.
     *
     * @author Nick Burch
     */

    public abstract class PositionDependentRecordAtom : RecordAtom, PositionDependentRecord
    {
        /** Our location on the disk, as of the last write out */
        protected int myLastOnDiskOffset;

        /** Fetch our location on the disk, as of the last write out */
        public int LastOnDiskOffset
        {
            get { return myLastOnDiskOffset; }
            set { myLastOnDiskOffset = value; }
        }
        /**
         * Offer the record the list of records that have Changed their
         *  location as part of the Writeout.
         * Allows records to update their internal pointers to other records
         *  locations
         */
        public abstract void UpdateOtherRecordReferences(Hashtable oldToNewReferencesLookup);
    }
}


