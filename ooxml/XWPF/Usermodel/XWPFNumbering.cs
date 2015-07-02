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
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.IO;
    using System.Xml.Serialization;
    using System.Xml;


    /**
     * @author Philipp Epp
     *
     */
    public class XWPFNumbering : POIXMLDocumentPart
    {
        protected List<XWPFAbstractNum> abstractNums = new List<XWPFAbstractNum>();
        protected List<XWPFNum> nums = new List<XWPFNum>();

        private CT_Numbering ctNumbering;
        bool isNew;

        /**
         *create a new styles object with an existing document 
         */
        public XWPFNumbering(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            isNew = true;
        }

        /**
         * create a new XWPFNumbering object for use in a new document
         */
        public XWPFNumbering()
        {
            abstractNums = new List<XWPFAbstractNum>();
            nums = new List<XWPFNum>();
            isNew = true;
        }

        /**
         * read numbering form an existing package
         */

        internal override void OnDocumentRead()
        {
            NumberingDocument numberingDoc = null;
            
            XmlDocument doc = ConvertStreamToXml(GetPackagePart().GetInputStream());
            try {
                numberingDoc = NumberingDocument.Parse(doc, NamespaceManager);
                ctNumbering = numberingDoc.Numbering;
                //get any Nums
                foreach(CT_Num ctNum in ctNumbering.GetNumList()) {
                    nums.Add(new XWPFNum(ctNum, this));
                }
                foreach(CT_AbstractNum ctAbstractNum in ctNumbering.GetAbstractNumList()){
                    abstractNums.Add(new XWPFAbstractNum(ctAbstractNum, this));
                }
                isNew = false;
            } catch (Exception e) {
                throw new POIXMLException(e);
            }
        }

        /**
         * save and Commit numbering
         */

        protected internal override void Commit()
        {
            /*XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.SaveSyntheticDocumentElement=(new QName(CTNumbering.type.Name.NamespaceURI, "numbering"));
            Dictionary<String,String> map = new Dictionary<String,String>();
            map.Put("http://schemas.Openxmlformats.org/markup-compatibility/2006", "ve");
            map.Put("urn:schemas-microsoft-com:office:office", "o");
            map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/relationships", "r");
            map.Put("http://schemas.Openxmlformats.org/officeDocument/2006/math", "m");
            map.Put("urn:schemas-microsoft-com:vml", "v");
            map.Put("http://schemas.Openxmlformats.org/drawingml/2006/wordProcessingDrawing", "wp");
            map.Put("urn:schemas-microsoft-com:office:word", "w10");
            map.Put("http://schemas.Openxmlformats.org/wordProcessingml/2006/main", "w");
            map.Put("http://schemas.microsoft.com/office/word/2006/wordml", "wne");
            xmlOptions.SaveSuggestedPrefixes=(map);*/
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            NumberingDocument doc = new NumberingDocument(ctNumbering);
            //XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            //    new XmlQualifiedName("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
            //    new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            //    new XmlQualifiedName("m", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
            //    new XmlQualifiedName("v", "urn:schemas-microsoft-com:vml"),
            //    new XmlQualifiedName("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
            //    new XmlQualifiedName("w10", "urn:schemas-microsoft-com:office:word"),
            //    new XmlQualifiedName("wne", "http://schemas.microsoft.com/office/word/2006/wordml"),
            //     new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            // });
            doc.Save(out1);
            out1.Close();
        }


        /**
         * Sets the ctNumbering
         * @param numbering
         */
        public void SetNumbering(CT_Numbering numbering)
        {
            ctNumbering = numbering;
        }


        /**
         * Checks whether number with numID exists
         * @param numID
         * @return bool		true if num exist, false if num not exist
         */
        public bool NumExist(string numID)
        {
            foreach (XWPFNum num in nums)
            {
                if (num.GetCTNum().numId.Equals(numID))
                    return true;
            }
            return false;
        }

        /**
         * add a new number to the numbering document
         * @param num
         */
        public string AddNum(XWPFNum num){
            ctNumbering.AddNewNum();
            int pos = (ctNumbering.GetNumList().Count) - 1;
            ctNumbering.SetNumArray(pos, num.GetCTNum());
            nums.Add(num);
            return num.GetCTNum().numId;
        }

        /**
         * Add a new num with an abstractNumID
         * @return return NumId of the Added num 
         */
        public string AddNum(string abstractNumID)
        {
            CT_Num ctNum = this.ctNumbering.AddNewNum();
            ctNum.AddNewAbstractNumId();
            ctNum.abstractNumId.val = (abstractNumID);
            ctNum.numId = (nums.Count + 1).ToString();
            XWPFNum num = new XWPFNum(ctNum, this);
            nums.Add(num);
            return ctNum.numId;
        }

        /**
         * Add a new num with an abstractNumID and a numID
         * @param abstractNumID
         * @param numID
         */
        public void AddNum(string abstractNumID, string numID)
        {
            CT_Num ctNum = this.ctNumbering.AddNewNum();
            ctNum.AddNewAbstractNumId();
            ctNum.abstractNumId.val = (abstractNumID);
            ctNum.numId = (numID);
            XWPFNum num = new XWPFNum(ctNum, this);
            nums.Add(num);
        }

        /**
         * Get Num by NumID
         * @param numID
         * @return abstractNum with NumId if no Num exists with that NumID 
         * 			null will be returned
         */
        public XWPFNum GetNum(string numID){
            foreach(XWPFNum num in nums){
                if(num.GetCTNum().numId.Equals(numID))
                    return num;
            }
            return null;
        }
        /**
         * Get AbstractNum by abstractNumID
         * @param abstractNumID
         * @return  abstractNum with abstractNumId if no abstractNum exists with that abstractNumID 
         * 			null will be returned
         */
        public XWPFAbstractNum GetAbstractNum(string abstractNumID){
            foreach(XWPFAbstractNum abstractNum in abstractNums){
                if(abstractNum.GetAbstractNum().abstractNumId.Equals(abstractNumID)){
                    return abstractNum;
                }
            }
            return null;
        }
        /**
         * Compare AbstractNum with abstractNums of this numbering document.
         * If the content of abstractNum Equals with an abstractNum of the List in numbering
         * the Bigint Value of it will be returned.
         * If no equal abstractNum is existing null will be returned
         * 
         * @param abstractNum
         * @return 	Bigint
         */
        public string GetIdOfAbstractNum(XWPFAbstractNum abstractNum)
        {
            CT_AbstractNum copy = (CT_AbstractNum)abstractNum.GetCTAbstractNum().Copy();
            XWPFAbstractNum newAbstractNum = new XWPFAbstractNum(copy, this);
            int i;
            for (i = 0; i < abstractNums.Count; i++)
            {
                newAbstractNum.GetCTAbstractNum().abstractNumId = i.ToString();
                newAbstractNum.SetNumbering(this);
                if (newAbstractNum.GetCTAbstractNum().ValueEquals(abstractNums[i].GetCTAbstractNum()))
                {
                    return newAbstractNum.GetCTAbstractNum().abstractNumId;
                }
            }
            return null;
        }


        /**
         * add a new AbstractNum and return its AbstractNumID 
         * @param abstractNum
         */
        public string AddAbstractNum(XWPFAbstractNum abstractNum)
        {
            int pos = abstractNums.Count;
            if (abstractNum.GetAbstractNum() != null)
            { // Use the current CTAbstractNum if it exists
                ctNumbering.AddNewAbstractNum().Set(abstractNum.GetAbstractNum());
            }
            else
            {
                ctNumbering.AddNewAbstractNum();
                abstractNum.GetAbstractNum().abstractNumId = pos.ToString();
                ctNumbering.SetAbstractNumArray(pos, abstractNum.GetAbstractNum());
            }
            abstractNums.Add(abstractNum);
            return abstractNum.GetAbstractNum().abstractNumId;
        }
        /// <summary>
        /// Add a new AbstractNum
        /// </summary>
        /// <returns></returns>
        /// @author antony liu
        public string AddAbstractNum()
        {
            CT_AbstractNum ctAbstractNum = ctNumbering.AddNewAbstractNum();
            XWPFAbstractNum abstractNum = new XWPFAbstractNum(ctAbstractNum, this);
            abstractNum.AbstractNumId = abstractNums.Count.ToString();
            abstractNum.MultiLevelType = MultiLevelType.HybridMultilevel;
            abstractNum.InitLvl();
            abstractNums.Add(abstractNum);
            return abstractNum.GetAbstractNum().abstractNumId;
        }
        /**
         * remove an existing abstractNum 
         * @param abstractNumID
         * @return true if abstractNum with abstractNumID exists in NumberingArray,
         * 		   false if abstractNum with abstractNumID not exists
         */
        public bool RemoveAbstractNum(string abstractNumID)
        {
            if (int.Parse(abstractNumID) < abstractNums.Count)
            {
                ctNumbering.RemoveAbstractNum(int.Parse(abstractNumID));
                abstractNums.RemoveAt(int.Parse(abstractNumID));
                return true;
            }
            return false;
        }
        /**
         *return the abstractNumID
         *If the AbstractNumID not exists
         *return null
         * @param 		numID
         * @return 		abstractNumID
         */
        public string GetAbstractNumID(string numID)
        {
            XWPFNum num = GetNum(numID);
            if (num == null)
                return null;
            if (num.GetCTNum() == null)
                return null;
            if (num.GetCTNum().abstractNumId == null)
                return null;
            return num.GetCTNum().abstractNumId.val;
        }
    }

}
