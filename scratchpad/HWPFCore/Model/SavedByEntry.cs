/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace NPOI.HWPF.Model
{
    using System;


    /**
     * A single entry in the {@link SavedByTable}.
     * 
     * @author Daniel Noll
     */
    public class SavedByEntry
    {
        private String userName;
        private String saveLocation;

        public SavedByEntry(String userName, String saveLocation)
        {
            this.userName = userName;
            this.saveLocation = saveLocation;
        }

        public String GetUserName()
        {
            return userName;
        }

        public String GetSaveLocation()
        {
            return saveLocation;
        }

        /**
         * Compares this object with another, for equality.
         *
         * @param other the object to compare to this one.
         * @return <code>true</code> iff the other object Is equal to this one.
         */
        public override bool Equals(Object other)
        {
            if (other == this) return true;
            if (!(other is SavedByEntry)) return false;
            SavedByEntry that = (SavedByEntry)other;
            return that.userName.Equals(userName) &&
                   that.saveLocation.Equals(saveLocation);
        }

        /**
         * Generates a hash code for consistency with {@link #equals(Object)}.
         *
         * @return the hash code.
         */
        public override int GetHashCode()
        {
            int hash = 29;
            hash = hash * 13 + userName.GetHashCode();
            hash = hash * 13 + saveLocation.GetHashCode();
            return hash;
        }

        /**
         * Returns a string for display.
         *
         * @return the string.
         */
        public override String ToString()
        {
            return "SavedByEntry[userName=" + GetUserName() +
                               ",saveLocation=" + GetSaveLocation() + "]";
        }
    }
}