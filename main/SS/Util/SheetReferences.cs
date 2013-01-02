/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


using System;
using System.Collections;


namespace NPOI.SS.Util
{

/**
 * Holds a collection of Sheet names and their associated
 * reference numbers.
 *
 * @author Andrew C. Oliver (acoliver at apache dot org)
 *
 */
public class SheetReferences
{
    Hashtable map;
    public SheetReferences()
    {
        map = new Hashtable(5);
    }
 
    public void AddSheetReference(String sheetName, int number) {
       map[number]= sheetName;
    } 

    public String GetSheetName(int number) {
       return (String)map[number];
    }

}
}