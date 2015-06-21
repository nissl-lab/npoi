/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.OpenXmlFormats.Dml;
namespace NPOI.XSSF.Model
{
    /**
     *  Used internally to navigate the paragraph text style hierarchy within a shape and fetch properties
    */
    public abstract class ParagraphPropertyFetcher
    {
        public abstract bool Fetch(CT_TextParagraphProperties props);
        public abstract bool Fetch(CT_Shape props);
    }
    public abstract class ParagraphPropertyFetcher<T> : ParagraphPropertyFetcher
    {

        private T _value;
        private int _level;

        public T GetValue()
        {
            return _value;
        }

        public void SetValue(T val)
        {
            _value = val;
        }

        public ParagraphPropertyFetcher(int level)
        {
            _level = level;
        }

        /**
        *
        * @param shape the shape being examined
        * @return true if the desired property was fetched
        */
        public override bool Fetch(CT_Shape shape)
        {
            //XmlObject[] o = shape.selectPath(
            //        "declare namespace xdr='http://schemas.Openxmlformats.org/drawingml/2006/spreadsheetDrawing' " +
            //        "declare namespace a='http://schemas.Openxmlformats.org/drawingml/2006/main' " +
            //        ".//xdr:txBody/a:lstStyle/a:lvl" + (_level + 1) + "pPr"
            //);
            //if (o.Length == 1)
            //{
            //    CT_TextParagraphProperties props = (CT_TextParagraphProperties)o[0];
            //    return Fetch(props);
            //}
            return false;
        }

        
    }
}