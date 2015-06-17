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
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml.Picture;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XWPF.UserModel
{

    /**
     * @author Philipp Epp
     */
    public class XWPFPicture
    {

        private CT_Picture ctPic;
        private String description;
        private XWPFRun run;

        public XWPFPicture(CT_Picture ctPic, XWPFRun Run)
        {
            this.run = Run;
            this.ctPic = ctPic;
            description = ctPic.nvPicPr.cNvPr.descr;
        }

        /**
         * Link Picture with PictureData
         * @param rel
         */
        public void SetPictureReference(PackageRelationship rel)
        {
            ctPic.blipFill.blip.embed = (rel.Id);
        }

        /**
         * Return the underlying CTPicture bean that holds all properties for this picture
         *
         * @return the underlying CTPicture bean
         */
        public CT_Picture GetCTPicture()
        {
            return ctPic;
        }

        /**
         * Get the PictureData of the Picture, if present.
         * Note - not all kinds of picture have data
         */
        public XWPFPictureData GetPictureData()
        {
            //String blipId = ctPic.blipFill.blip.embed;
            CT_BlipFillProperties blipProps = ctPic.blipFill;

            if (blipProps == null || !blipProps.IsSetBlip())
            {
                // return null if Blip data is missing
                return null;
            }

            String blipId = blipProps.blip.embed;


            POIXMLDocumentPart part = run.Parent.Part;
            if (part != null)
            {
                POIXMLDocumentPart relatedPart = part.GetRelationById(blipId);
                if (relatedPart is XWPFPictureData)
                {
                    return (XWPFPictureData)relatedPart;
                }
            }
            return null;
        }

        public String GetDescription()
        {
            return description;
        }
    }

}