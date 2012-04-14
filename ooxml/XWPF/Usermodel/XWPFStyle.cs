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

    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * @author Philipp Epp
     *
     */
    public class XWPFStyle
    {

        private CT_Style ctStyle;
        protected XWPFStyles styles;

        /**
         * constructor
         * @param style
         */
        public XWPFStyle(CT_Style style)
            : this(style, null)
        {
        }
        /**
         * constructor
         * @param style
         * @param styles
         */
        public XWPFStyle(CT_Style style, XWPFStyles styles)
        {
            this.ctStyle = style;
            this.styles = styles;
        }

        /**
         * Get StyleID of the style
         * @return styleID		StyleID of the style
         */
        public String GetStyleId()
        {
            //return ctStyle.StyleId;
            throw new NotImplementedException();
        }

        /**
         * Get Type of the Style
         * @return	ctType 
         */
        public ST_StyleType GetType()
        {
            //return ctStyle.Type;
            throw new NotImplementedException();
        }

        /**
         * Set style
         * @param style		
         */
        public void SetStyle(CT_Style style)
        {
            this.ctStyle = style;
        }
        /**
         * Get ctStyle
         * @return	ctStyle
         */
        public CT_Style GetCTStyle()
        {
            return this.ctStyle;
        }
        /**
         * Set styleID
         * @param styleId
         */
        public void SetStyleId(String styleId)
        {
            //ctStyle.StyleId=(styleId);
            throw new NotImplementedException();
        }

        /**
         * Set styleType
         * @param type
         */
        public void SetType(ST_StyleType type)
        {
            //ctStyle.Type=(type);
            throw new NotImplementedException();
        }
        /**
         * Get styles
         * @return styles		the styles to which this style belongs
         */
        public XWPFStyles GetStyles()
        {
            return styles;
        }

        public String GetBasisStyleID()
        {
            /*if(ctStyle.BasedOn!=null)
                return ctStyle.BasedOn.Val;
            else
                return null;*/
            throw new NotImplementedException();
        }


        /**
         * Get StyleID of the linked Style
         */
        public String GetLinkStyleID()
        {
            /*if (ctStyle.Link!=null)
                return ctStyle.Link.Val;
            else
                return null;
             * */
            throw new NotImplementedException();
        }

        /**
         * Get StyleID of the next style
         */
        public String GetNextStyleID()
        {
            /*if(ctStyle.Next!=null)
                return ctStyle.Next.Val;
            else
                return null;
             * */
            throw new NotImplementedException();
        }

        public String GetName()
        {
            //if(ctStyle.IsSetName()) 
            //   return ctStyle.Name.Val;
            //return null;
            throw new NotImplementedException();
        }

        /**
         * Compares the names of the Styles 
         * @param compStyle
         */
        public bool HasSameName(XWPFStyle compStyle)
        {
            //CTStyle ctCompStyle = compStyle.CTStyle;
            //String name = ctCompStyle.Name.Val;
            //return name.Equals(ctStyle.Name.Val);
            throw new NotImplementedException();
        }

    }//end class

}