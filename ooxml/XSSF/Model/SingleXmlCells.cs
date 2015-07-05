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

using System.Xml;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.UserModel.Helpers;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
namespace NPOI.XSSF.Model
{


    /**
     * 
     * This class : the Single Cell Tables Part (Open Office XML Part 4:
     * chapter 3.5.2)
     * 
     *
     * @author Roberto Manicardi
     */
    public class SingleXmlCells : POIXMLDocumentPart
    {


        private CT_SingleXmlCells SingleXMLCells;

        public SingleXmlCells()
            : base()
        {

            SingleXMLCells = new CT_SingleXmlCells();
        }

        internal SingleXmlCells(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            ReadFrom(part.GetInputStream());
        }

        public void ReadFrom(Stream is1)
        {
            try
            {
                SingleXmlCellsDocument doc = SingleXmlCellsDocument.Parse(is1);
                SingleXMLCells = doc.GetSingleXmlCells();
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }

        public XSSFSheet GetXSSFSheet()
        {
            return (XSSFSheet)GetParent();
        }

        protected void WriteTo(Stream out1)
        {
            SingleXmlCellsDocument doc = new SingleXmlCellsDocument();
            doc.SetSingleXmlCells(SingleXMLCells);
            doc.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }

        public CT_SingleXmlCells GetCTSingleXMLCells()
        {
            return SingleXMLCells;
        }

        /**
         * 
         * @return all the SimpleXmlCell Contained in this SingleXmlCells element
         */
        public List<XSSFSingleXmlCell> GetAllSimpleXmlCell()
        {
            List<XSSFSingleXmlCell> list = new List<XSSFSingleXmlCell>();

            foreach (CT_SingleXmlCell SingleXmlCell in SingleXMLCells.singleXmlCell)
            {
                list.Add(new XSSFSingleXmlCell(SingleXmlCell, this));
            }
            return list;
        }
    }
}






