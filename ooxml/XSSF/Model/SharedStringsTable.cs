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


namespace NPOI.XSSF.Model
{
    using System.Collections.Generic;
    using OpenXmlFormats.Spreadsheet;
    using System;
    using System.IO;
    using NPOI.OpenXml4Net.OPC;
    using System.Xml;
    using System.Security;
    using System.Text.RegularExpressions;
    using System.Text;

    /**
     * Table of strings shared across all sheets in a workbook.
     * <p>
     * A workbook may contain thousands of cells Containing string (non-numeric) data. Furthermore this data is very
     * likely to be repeated across many rows or columns. The goal of implementing a single string table that is shared
     * across the workbook is to improve performance in opening and saving the file by only Reading and writing the
     * repetitive information once.
     * </p>
     * <p>
     * Consider for example a workbook summarizing information for cities within various countries. There may be a
     * column for the name of the country, a column for the name of each city in that country, and a column
     * Containing the data for each city. In this case the country name is repetitive, being duplicated in many cells.
     * In many cases the repetition is extensive, and a tremendous savings is realized by making use of a shared string
     * table when saving the workbook. When displaying text in the spreadsheet, the cell table will just contain an
     * index into the string table as the value of a cell, instead of the full string.
     * </p>
     * <p>
     * The shared string table Contains all the necessary information for displaying the string: the text, formatting
     * properties, and phonetic properties (for East Asian languages).
     * </p>
     *
     * @author Nick Birch
     * @author Yegor Kozlov
     */
    public class SharedStringsTable : POIXMLDocumentPart
    {

        /**
         *  Array of individual string items in the Shared String table.
         */
        private List<CT_Rst> strings = new List<CT_Rst>();

        /**
         *  Maps strings and their indexes in the <code>strings</code> arrays
         */
        private Dictionary<String, int> stmap = new Dictionary<String, int>();

        /**
         * An integer representing the total count of strings in the workbook. This count does not
         * include any numbers, it counts only the total of text strings in the workbook.
         */
        private int count;

        /**
         * An integer representing the total count of unique strings in the Shared String Table.
         * A string is unique even if it is a copy of another string, but has different formatting applied
         * at the character level.
         */
        private int uniqueCount;

        private SstDocument _sstDoc;

        public SharedStringsTable()
            : base()
        {

            _sstDoc = new SstDocument();
            _sstDoc.AddNewSst();
        }

        internal SharedStringsTable(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            ReadFrom(xml);
        }



        public void ReadFrom(XmlDocument xml)
        {
                 int cnt = 0;
                _sstDoc = SstDocument.Parse(xml, NamespaceManager);
                CT_Sst sst = _sstDoc.GetSst();
                count = (int)sst.count;
                uniqueCount = (int)sst.uniqueCount;
                foreach (CT_Rst st in sst.si)
                {
                     string key=GetKey(st);
                   if(key!=null && !stmap.ContainsKey(key))
                       stmap.Add(key, cnt);
                   strings.Add(st);
                    cnt++;
                }

        }

        private String GetKey(CT_Rst st)
        {
            return st.XmlText;
        }

        /**
         * Return a string item by index
         *
         * @param idx index of item to return.
         * @return the item at the specified position in this Shared String table.
         */
        public CT_Rst GetEntryAt(int idx)
        {
            return strings[idx];
        }

        /**
         * Return an integer representing the total count of strings in the workbook. This count does not
         * include any numbers, it counts only the total of text strings in the workbook.
         *
         * @return the total count of strings in the workbook
         */
        public int Count
        {
            get
            {
                return count;
            }
        }

        /**
         * Returns an integer representing the total count of unique strings in the Shared String Table.
         * A string is unique even if it is a copy of another string, but has different formatting applied
         * at the character level.
         *
         * @return the total count of unique strings in the workbook
         */
        public int UniqueCount
        {
            get
            {
                return uniqueCount;
            }
        }

        /**
         * Add an entry to this Shared String table (a new value is appened to the end).
         *
         * <p>
         * If the Shared String table already Contains this <code>CT_Rst</code> bean, its index is returned.
         * Otherwise a new entry is aded.
         * </p>
         *
         * @param st the entry to add
         * @return index the index of Added entry
         */
        public int AddEntry(CT_Rst st)
        {
            String s = GetKey(st);
            count++;
            if (stmap.ContainsKey(s))
            {
                return stmap[s];
            }

            uniqueCount++;
            //create a CT_Rst bean attached to this SstDocument and copy the argument CT_Rst into it
            CT_Rst newSt = new CT_Rst();
            _sstDoc.GetSst().si.Add(newSt);
            newSt.Set(st);
            int idx = strings.Count;
            stmap[s] = idx;
            strings.Add(newSt);
            return idx;
        }
        /**
         * Provide low-level access to the underlying array of CT_Rst beans
         *
         * @return array of CT_Rst beans
         */
        public List<CT_Rst> Items
        {
            get
            {
                return strings;
            }
        }

        /**
         * 
         * this table out as XML.
         * 
         * @param out The stream to write to.
         * @throws IOException if an error occurs while writing.
         */
        public void WriteTo(Stream out1)
        {
            // the following two lines turn off writing CDATA
            // see Bugzilla 48936
            //options.SetSaveCDataLengthThreshold(1000000);
            //options.SetSaveCDataEntityCountThreshold(-1);
            CT_Sst sst = _sstDoc.GetSst();
            sst.count = count;
           sst.uniqueCount = uniqueCount;

           //re-create the sst table every time saving a workbook
           _sstDoc.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            //Stream out1 = part.GetInputStream();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }
    }
}






