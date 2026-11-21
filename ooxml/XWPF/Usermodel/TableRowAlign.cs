/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XWPF.UserModel
{
    /// <summary>
    /// Sets alignment values allowed for Tables and Table Rows
    /// </summary>
    public enum TableRowAlign
    {

        LEFT = 0,
        CENTER = 1,
        RIGHT = 2
    }
    public static class TableRowAlignExtension
    {
        private static readonly Dictionary<int, TableRowAlign> imap = 
            new(){
                { 0, TableRowAlign.LEFT },
                { 1, TableRowAlign.CENTER },
                { 2, TableRowAlign.RIGHT },
            };

        public static TableRowAlign ValueOf(int type)
        {
            if(imap.TryGetValue(type, out TableRowAlign err))
                return err;
            throw new ArgumentException("Unknown table row alignment: " + type);
        }

        public static int GetValue(this TableRowAlign trAlign)
        {
            return (int) trAlign;
        }
    }
}


