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

using System;
using System.Collections.Generic;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Xml;

namespace NPOI.XSSF.Model
{

    /**
     * The cells in a workbook can be calculated in different orders depending on various optimizations and
     * dependencies. The calculation chain object specifies the order in which the cells in a workbook were last calculated.
     *
     * @author Yegor Kozlov
     */
    public class CalculationChain : POIXMLDocumentPart
    {
        private CT_CalcChain chain;

        public CalculationChain()
            : base()
        {

            chain = new CT_CalcChain();
        }

        internal CalculationChain(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            ReadFrom(xml);
        }

        public void ReadFrom(XmlDocument xml)
        {
            CalcChainDocument doc = CalcChainDocument.Parse(xml, NamespaceManager);
            chain = doc.GetCalcChain();

        }
        public void WriteTo(Stream out1)
        {
            CalcChainDocument doc = new CalcChainDocument();
            doc.SetCalcChain(chain);
            doc.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }


        public CT_CalcChain GetCTCalcChain()
        {
            return chain;
        }

        /**
         * Remove a formula reference from the calculation chain
         * 
         * @param sheetId  the sheet Id of a sheet the formula belongs to.
         * @param ref  A1 style reference to the cell Containing the formula.
         */
        //  GetXYZArray() array accessors are deprecated 
        public void RemoveItem(int sheetId, String ref1)
        {
            //sheet Id of a sheet the cell belongs to
            int id = -1;
            List<CT_CalcCell> c = chain.c;

            for (int i = 0; i < c.Count; i++)
            {
                //If sheet Id  is omitted, it is assumed to be the same as the value of the previous cell.
                if (c[i].iSpecified)
                {
                    id = c[i].i;
                }

                if (id == sheetId && c[i].r.Equals(ref1))
                {
                    //if (c[i].IsSetI() && i < c.Length - 1 && !c[i + 1].IsSetI())
                    //if (i < c.Count - 1)
                    if (c[i].iSpecified && i < c.Count - 1 && !c[i + 1].iSpecified)
                    {
                        c[i + 1].i = id;
                    }
                    chain.RemoveC(i);
                    break;
                }
            }
        }
    }
}
