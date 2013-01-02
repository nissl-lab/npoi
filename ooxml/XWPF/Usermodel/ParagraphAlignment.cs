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
     * Specifies all types of alignment which are available to be applied to objects in a
     * WordProcessingML document
     *
     * @author Yegor Kozlov
     */
    public enum ParagraphAlignment
    {
        //YK: TODO document each alignment option

        LEFT = (1),
        CENTER = (2),
        RIGHT = (3),
        BOTH = (4),
        MEDIUM_KASHIDA = (5),
        DISTRIBUTE = (6),
        NUM_TAB = (7),
        HIGH_KASHIDA = (8),
        LOW_KASHIDA = (9),
        THAI_DISTRIBUTE = (10)

        //private int value;

        //private ParagraphAlignment(int val){
        //    value = val;
        //}

        //public int GetValue(){
        //    return value;
        //}

        //private static Dictionary<int, ParagraphAlignment> imap = new Dictionary<int, ParagraphAlignment>();
        //static{
        //    foreach (ParagraphAlignment p in values()) {
        //        imap.Put(new int(p.Value), p);
        //    }
        //}

        //public static ParagraphAlignment ValueOf(int type){
        //    ParagraphAlignment err = imap.Get(new int(type));
        //    if(err == null) throw new ArgumentException("Unknown paragraph alignment: " + type);
        //    return err;
        //}

    }
}