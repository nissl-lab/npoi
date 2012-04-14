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


    /**
     * @author Philipp Epp
     *
     */
    public class XWPFNumbering : POIXMLDocumentPart
    {
        //protected List<XWPFAbstractNum> abstractNums = new List<XWPFAbstractNum>();
        //protected List<XWPFNum> nums = new List<XWPFNum>();

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
            //abstractNums = new List<XWPFAbstractNum>();
            //nums = new List<XWPFNum>();
            //isNew = true;
        }

        /**
         * read numbering form an existing package
         */

        protected void onDocumentRead()
        {
            /*NumberingDocument numberingDoc = null;
            Stream is1;
            is1 = GetPackagePart().GetInputStream();
            try {
                numberingDoc = NumberingDocument.Factory.Parse(is1);
                ctNumbering = numberingDoc.Numbering;
                //get any Nums
                foreach(CT_Num ctNum in ctNumbering.NumList) {
                    nums.Add(new XWPFNum(ctNum, this));
                }
                foreach(CT_AbstractNum ctAbstractNum in ctNumbering.AbstractNumList){
                    abstractNums.Add(new XWPFAbstractNum(ctAbstractNum, this));
                }
                isNew = false;
            } catch (Exception e) {
                throw new POIXMLException();
            }*/
        }

        /**
         * save and Commit numbering
         */

        protected void Commit()
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
            xmlOptions.SaveSuggestedPrefixes=(map);
            PackagePart part = GetPackagePart();
            OutputStream out1 = part.OutputStream;
            ctNumbering.Save(out, xmlOptions);
            out1.Close();*/
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
        //public bool numExist(Bigint numID){
        //    foreach (XWPFNum num in nums) {
        //        if (num.GetCTNum().NumId.Equals(numID))
        //            return true;
        //    }
        //    return false;
        //}

        /**
         * add a new number to the numbering document
         * @param num
         */
        //public Bigint AddNum(XWPFNum num){
        //    ctNumbering.AddNewNum();
        //    int pos = (ctNumbering.NumList.Size()) - 1;
        //    ctNumbering.NumArray=(pos, num.CTNum);
        //    nums.Add(num);
        //    return num.CTNum.NumId;
        //}

        /**
         * Add a new num with an abstractNumID
         * @return return NumId of the Added num 
         */
        //public Bigint AddNum(Bigint abstractNumID){
        //    CTNum ctNum = this.ctNumbering.AddNewNum();
        //    ctNum.AddNewAbstractNumId();
        //    ctNum.AbstractNumId.Val=(abstractNumID);
        //    ctNum.NumId=(BigInt32.ValueOf(nums.Size()+1));
        //    XWPFNum num = new XWPFNum(ctNum, this);
        //    nums.Add(num);
        //    return ctNum.NumId;
        //}

        /**
         * Add a new num with an abstractNumID and a numID
         * @param abstractNumID
         * @param numID
         */
        //public void AddNum(Bigint abstractNumID, Bigint numID){
        //    CTNum ctNum = this.ctNumbering.AddNewNum();
        //    ctNum.AddNewAbstractNumId();
        //    ctNum.AbstractNumId.Val=(abstractNumID);
        //    ctNum.NumId=(numID);
        //    XWPFNum num = new XWPFNum(ctNum, this);
        //    nums.Add(num);
        //}

        /**
         * Get Num by NumID
         * @param numID
         * @return abstractNum with NumId if no Num exists with that NumID 
         * 			null will be returned
         */
        //public XWPFNum GetNum(Bigint numID){
        //    foreach(XWPFNum num: nums){
        //        if(num.CTNum.NumId.Equals(numID))
        //            return num;
        //    }
        //    return null;
        //}
        /**
         * Get AbstractNum by abstractNumID
         * @param abstractNumID
         * @return  abstractNum with abstractNumId if no abstractNum exists with that abstractNumID 
         * 			null will be returned
         */
        //public XWPFAbstractNum GetAbstractNum(Bigint abstractNumID){
        //    foreach(XWPFAbstractNum abstractNum: abstractNums){
        //        if(abstractNum.AbstractNum.AbstractNumId.Equals(abstractNumID)){
        //            return abstractNum;
        //        }
        //    }
        //    return null;
        //}
        /**
         * Compare AbstractNum with abstractNums of this numbering document.
         * If the content of abstractNum Equals with an abstractNum of the List in numbering
         * the Bigint Value of it will be returned.
         * If no equal abstractNum is existing null will be returned
         * 
         * @param abstractNum
         * @return 	Bigint
         */
        //public Bigint GetIdOfAbstractNum(XWPFAbstractNum abstractNum){
        //    CTAbstractNum copy = (CTAbstractNum) abstractNum.CTAbstractNum.Copy();
        //    XWPFAbstractNum newAbstractNum = new XWPFAbstractNum(copy, this);
        //    int i;
        //    for (i = 0; i < abstractNums.Size(); i++) {
        //        newAbstractNum.CTAbstractNum.AbstractNumId=(BigInt32.ValueOf(i));
        //        newAbstractNum.Numbering=(this);
        //        if(newAbstractNum.CTAbstractNum.ValueEquals(abstractNums.Get(i).CTAbstractNum)){
        //            return newAbstractNum.CTAbstractNum.AbstractNumId;
        //        }
        //    }
        //    return null;
        //}


        /**
         * add a new AbstractNum and return its AbstractNumID 
         * @param abstractNum
         */
        //public Bigint AddAbstractNum(XWPFAbstractNum abstractNum){
        //    int pos = abstractNums.Size();
        //    if(abstractNum.AbstractNum != null){ // Use the current CTAbstractNum if it exists
        //        ctNumbering.AddNewAbstractNum().Set(abstractNum.AbstractNum);
        //    } else {
        //        ctNumbering.AddNewAbstractNum();
        //        abstractNum.AbstractNum.AbstractNumId=(BigInt32.ValueOf(pos));
        //        ctNumbering.AbstractNumArray=(pos, abstractNum.AbstractNum);
        //    }
        //    abstractNums.Add(abstractNum);
        //    return abstractNum.CTAbstractNum.AbstractNumId;
        //}

        /**
         * remove an existing abstractNum 
         * @param abstractNumID
         * @return true if abstractNum with abstractNumID exists in NumberingArray,
         * 		   false if abstractNum with abstractNumID not exists
         */
        //public bool RemoveAbstractNum(Bigint abstractNumID){
        //    if(abstractNumID.ByteValue()<abstractNums.Size()){
        //        ctNumbering.RemoveAbstractNum(abstractNumID.ByteValue());
        //        abstractNums.Remove(abstractNumID.ByteValue());
        //        return true;
        //    }
        //    return false;
        //}
        /**
         *return the abstractNumID
         *If the AbstractNumID not exists
         *return null
         * @param 		numID
         * @return 		abstractNumID
         */
        //public Bigint GetAbstractNumID(Bigint numID){
        //    XWPFNum num = GetNum(numID);
        //    if(num == null)
        //        return null;
        //    if (num.CTNum == null)
        //        return null;
        //    if (num.CTNum.AbstractNumId == null)
        //        return null;
        //    return num.CTNum.AbstractNumId.Val;
        //}
    }

}
