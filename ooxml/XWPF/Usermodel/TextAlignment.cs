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
     * Specifies all types of vertical alignment which are available to be applied to of all text 
     * on each line displayed within a paragraph.
     * 
     * @author Gisella Bronzetti
     */
    public enum TextAlignment
    {
        /**
         * Specifies that all text in the parent object shall be 
         * aligned to the top of each character when displayed
         */
        TOP = (1),
        /**
         * Specifies that all text in the parent object shall be 
         * aligned to the center of each character when displayed.
         */
        CENTER = (2),
        /**
         * Specifies that all text in the parent object shall be
         * aligned to the baseline of each character when displayed.
         */
        BASELINE = (3),
        /**
         * Specifies that all text in the parent object shall be
         * aligned to the bottom of each character when displayed.
         */
        BOTTOM = (4),
        /**
         * Specifies that all text in the parent object shall be 
         * aligned automatically when displayed.
         */
        AUTO = (5)

        //private int value;

        //private TextAlignment(int val){
        //value = val;
        //}

        //public int GetValue(){
        //   return value;
        //}

        //private static Dictionary<int, TextAlignment> imap = new Dictionary<int, TextAlignment>();
        //static{
        //   foreach (TextAlignment p in values()) {
        //      imap.Put(new int(p.Value), p);
        //   }
        //}

        //public static TextAlignment ValueOf(int type){
        //   TextAlignment align = imap.Get(new int(type));
        //   if(align == null) throw new ArgumentException("Unknown text alignment: " + type);
        //   return align;
        //}
    }
}

