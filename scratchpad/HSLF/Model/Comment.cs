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

namespace NPOI.HSLF.Model
{
    using NPOI.HSLF.Record;
    using System;

    /**
     *
     * @author Nick Burch
     */
    public class Comment
    {
        private Comment2000 _comment2000;

        public Comment(Comment2000 comment2000)
        {
            _comment2000 = comment2000;
        }

        protected Comment2000 GetComment2000()
        {
            return _comment2000;
        }

        /**
         * Get the Author of this comment
         */
        public String Author
        {
            get
            {
                return _comment2000.Author;
            }
            set
            {
                _comment2000.Author = value;
            }
        }

        /**
         * Get the Author's Initials of this comment
         */
        public String AuthorInitials
        {
            get
            {
                return _comment2000.AuthorInitials;
            }
            set
            {
                _comment2000.AuthorInitials = value;
            }
        }
        /**
         * Get the text of this comment
         */
        public String Text
        {
            get
            {
                return _comment2000.Text;
            }
            set
            {
                _comment2000.Text = value;
            }
        }
    }



}

