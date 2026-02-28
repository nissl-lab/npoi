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
     * Specifies the types of patterns which may be used to create the underline
     * applied beneath the text in a Run.
     * 
     * @author Gisella Bronzetti
     */
    public enum UnderlinePatterns
    {

        /**
         * Specifies an underline consisting of a single line beneath all characters
         * in this Run.
         */
        Single = (1),

        /**
         * Specifies an underline consisting of a single line beneath all non-space
         * characters in the Run. There shall be no underline beneath any space
         * character (breaking or non-breaking).
         */
        Words = (2),

        /**
         * Specifies an underline consisting of two lines beneath all characters in
         * this run
         */
        Double = (3),

        /**
         * Specifies an underline consisting of a single thick line beneath all
         * characters in this Run.
         */
        Thick = (4),

        /**
         * Specifies an underline consisting of a series of dot characters beneath
         * all characters in this Run.
         */
        Dotted = (5),

        /**
         * Specifies an underline consisting of a series of thick dot characters
         * beneath all characters in this Run.
         */
        DottedHeavy = (6),

        /**
         * Specifies an underline consisting of a dashed line beneath all characters
         * in this Run.
         */
        Dash = (7),

        /**
         * Specifies an underline consisting of a series of thick dashes beneath all
         * characters in this Run.
         */
        DashedHeavy = (8),

        /**
         * Specifies an underline consisting of long dashed characters beneath all
         * characters in this Run.
         */
        DashLong = (9),

        /**
         * Specifies an underline consisting of thick long dashed characters beneath
         * all characters in this Run.
         */
        DashLongHeavy = (10),

        /**
         * Specifies an underline consisting of a series of dash, dot characters
         * beneath all characters in this Run.
         */
        DotDash = (11),

        /**
         * Specifies an underline consisting of a series of thick dash, dot
         * characters beneath all characters in this Run.
         */
        DashDotHeavy = (12),

        /**
         * Specifies an underline consisting of a series of dash, dot, dot
         * characters beneath all characters in this Run.
         */
        DotDotDash = (13),

        /**
         * Specifies an underline consisting of a series of thick dash, dot, dot
         * characters beneath all characters in this Run.
         */
        DashDotDotHeavy = (14),

        /**
         * Specifies an underline consisting of a single wavy line beneath all
         * characters in this Run.
         */
        Wave = (15),

        /**
         * Specifies an underline consisting of a single thick wavy line beneath all
         * characters in this Run.
         */
        WavyHeavy = (16),

        /**
         * Specifies an underline consisting of a pair of wavy lines beneath all
         * characters in this Run.
         */
        WavyDouble = (17),

        /**
         * Specifies no underline beneath this Run.
         */
        None = (18)

        //private int value;

        //private UnderlinePatterns(int val) {
        //   value = val;
        //}

        //public int GetValue() {
        //   return value;
        //}

        //private static Dictionary<int, UnderlinePatterns> imap = new Dictionary<int, UnderlinePatterns>();
        //static {
        //   foreach (UnderlinePatterns p in values()) {
        //      imap.Put(new int(p.Value), p);
        //   }
        //}

        //public static UnderlinePatterns ValueOf(int type) {
        //   UnderlinePatterns align = imap.Get(new int(type));
        //   if (align == null)
        //      throw new ArgumentException("Unknown underline pattern: "
        //            + type);
        //   return align;
        //}
    }

}