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
     * Specifies possible values for the alignment of the contents of this run in
     * relation to the default appearance of the Run's text. This allows the text to
     * be repositioned as subscript or superscript without altering the font size of
     * the run properties.
     * 
     * @author Gisella Bronzetti
     */
    public enum VerticalAlign
    {

        /**
         * Specifies that the text in the parent run shall be located at the
         * baseline and presented in the same size as surrounding text.
         */
        BASELINE = (1),
        /**
         * Specifies that this text should be subscript. This Setting shall lower
         * the text in this run below the baseline and change it to a smaller size,
         * if a smaller size is available.
         */
        SUPERSCRIPT = (2),
        /**
         * Specifies that this text should be superscript. This Setting shall raise
         * the text in this run above the baseline and change it to a smaller size,
         * if a smaller size is available.
         */
        SUBSCRIPT = (3)

        //private int value;

        //private VerticalAlign(int val) {
        //   value = val;
        //}

        //public int GetValue() {
        //   return value;
        //}

        //private static Dictionary<int, VerticalAlign> imap = new Dictionary<int, VerticalAlign>();
        //static {
        //   foreach (VerticalAlign p in values()) {
        //      imap.Put(new int(p.Value), p);
        //   }
        //}

        //public static VerticalAlign ValueOf(int type) {
        //   VerticalAlign align = imap.Get(new int(type));
        //   if (align == null)
        //      throw new ArgumentException("Unknown vertical alignment: "
        //            + type);
        //   return align;
        //}
    }
}

