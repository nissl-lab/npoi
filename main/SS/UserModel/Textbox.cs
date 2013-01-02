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

namespace NPOI.SS.UserModel
{

    public interface ITextbox : IShape
    {

        //public const short OBJECT_TYPE_TEXT = 6;

        /**
         * @return  the rich text string for this textbox.
         */
        IRichTextString String { get; set; }


        /**
         * @return  Returns the left margin within the textbox.
         */
        int MarginLeft { get; set; }


        /**
         * @return    returns the right margin within the textbox.
         */
        int MarginRight { get; set; }


        /**
         * @return  returns the top margin within the textbox.
         */
        int MarginTop { get; set; }

        /**
         * s the bottom margin within the textbox.
         */
        int MarginBottom { get; set; }

        short HorizontalAlignment { get; set; }
        short VerticalAlignment { get; set; }
    }
}