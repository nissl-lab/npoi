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
     * Specifies the Set of possible restart locations which may be used as to
     * determine the next available line when a break's type attribute has a value
     * of textWrapping.
     * 
     * @author Gisella Bronzetti
     */
    public enum BreakClear
    {

        /**
         * Specifies that the text wrapping break shall advance the text to the next
         * line in the WordProcessingML document, regardless of its position left to
         * right or the presence of any floating objects which intersect with the
         * line,
         * 
         * This is the Setting for a typical line break in a document.
         */

        NONE = (1),

        /**
         * Specifies that the text wrapping break shall behave as follows:
         * <ul>
         * <li> If this line is broken into multiple regions (a floating object in
         * the center of the page has text wrapping on both sides:
         * <ul>
         * <li> If this is the leftmost region of text flow on this line, advance
         * the text to the next position on the line </li>
         * <li>Otherwise, treat this as a text wrapping break of type all. </li>
         * </ul>
         * </li>
         * <li> If this line is not broken into multiple regions, then treat this
         * break as a text wrapping break of type none. </li>
         * </ul>
         * <li> If the parent paragraph is right to left, then these behaviors are
         * also reversed. </li>
         */
        LEFT = (2),

        /**
         * Specifies that the text wrapping break shall behave as follows:
         * <ul>
         * <li> If this line is broken into multiple regions (a floating object in
         * the center of the page has text wrapping on both sides:
         * <ul>
         * <li> If this is the rightmost region of text flow on this line, advance
         * the text to the next position on the next line </li>
         * <li> Otherwise, treat this as a text wrapping break of type all. </li>
         * </ul>
         * <li> If this line is not broken into multiple regions, then treat this
         * break as a text wrapping break of type none. </li>
         * <li> If the parent paragraph is right to left, then these beha viors are
         * also reversed. </li>
         * </ul>
         */
        RIGHT = (3),

        /**
         * Specifies that the text wrapping break shall advance the text to the next
         * line in the WordProcessingML document which spans the full width of the
         * line.
         */
        ALL = (4)

        //private int value;

        //private BreakClear(int val) {
        //   value = val;
        //}

        //public int GetValue() {
        //   return value;
        //}

        //private static Dictionary<int, BreakClear> imap = new Dictionary<int, BreakClear>();
        //static {
        //   foreach (BreakClear p in values()) {
        //      imap.Put(new int(p.Value), p);
        //   }
        //}

        //public static BreakClear ValueOf(int type) {
        //   BreakClear bType = imap.Get(new int(type));
        //   if (bType == null)
        //      throw new ArgumentException("Unknown break clear type: "
        //            + type);
        //   return bType;
        //}
    }

}