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

using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;

namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a text box in a SpreadsheetML Drawing.
     *
     * @author Yegor Kozlov
     */
    public class XSSFTextBox : XSSFSimpleShape //, ITextbox
    {

        internal XSSFTextBox(XSSFDrawing drawing, CT_Shape ctShape)
            : base(drawing, ctShape)
        {

        }

        //public short HorizontalAlignment
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public int MarginBottom
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public int MarginLeft
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public int MarginRight
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public int MarginTop
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public IRichTextString String
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public short VerticalAlignment
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public int CountOfAllChildren
        //{
        //    get { throw new System.NotImplementedException(); }
        //}

        //public int FillColor
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public override LineStyle LineStyle
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public override int LineStyleColor
        //{
        //    get { throw new System.NotImplementedException(); }
        //}

        //public override double LineWidth
        //{
        //    get
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new System.NotImplementedException();
        //    }
        //}

        //public void SetLineStyleColor(int lineStyleColor)
        //{
        //    throw new System.NotImplementedException();
        //}
    }
}


