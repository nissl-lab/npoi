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
namespace NPOI.XWPF.UserModel
{
    using System;




    /**
     * Specifies the possible types of break characters in a WordProcessingML
     * document.
     * The break type determines the next location where text shall be
     * placed After this manual break is applied to the text contents
     * 
     * @author Gisella Bronzetti
     */
    public enum BreakType
    {


        /**
         * Specifies that the current break shall restart itself on the next page of
         * the document when the document is displayed in page view.
         */
        PAGE = (1),

        /**
         * Specifies that the current break shall restart itself on the next column
         * available on the current page when the document is displayed in page
         * view.
         * <p>
         * If the current section is not divided into columns, or the column break
         * occurs in the last column on the current page when displayed, then the
         * restart location for text shall be the next page in the document.
         * </p>
         */
        COLUMN = (2),

        /**
         * Specifies that the current break shall restart itself on the next line in
         * the document when the document is displayed in page view.
         * The determine of the next line shall be done subject to the value of the clear
         * attribute on the specified break character.
         */
        TEXTWRAPPING = (3)

        //private int value;

        //private BreakType(int val) {
        //   value = val;
        //}

        //public int GetValue() {
        //   return value;
        //}

        //private static Dictionary<int, BreakType> imap = new Dictionary<int, BreakType>();
        //static {
        //   foreach (BreakType p in values()) {
        //      imap.Put(new int(p.Value), p);
        //   }
        //}

        //public static BreakType ValueOf(int type) {
        //   BreakType bType = imap.Get(new int(type));
        //   if (bType == null)
        //      throw new ArgumentException("Unknown break type: "
        //            + type);
        //   return bType;
        //}
    }

}