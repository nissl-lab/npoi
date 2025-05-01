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
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using EnumsNET;
using NPOI.OOXML.XSSF.UserModel;
namespace NPOI.XSSF.UserModel
{

/**
 * @author Yegor Kozlov
 */
    public class XSSFBorderFormatting : IBorderFormatting
    {
        IIndexedColorMap _colorMap;
        readonly CT_Border _border;

        /*package*/
        internal XSSFBorderFormatting(CT_Border border, IIndexedColorMap colorMap)
        {
            _border = border;
            _colorMap = colorMap;
        }

        #region IBorderFormatting Members

        public BorderStyle BorderBottom
        {
            get
            {
                return GetBorderStyle(_border.bottom);
            }
            set
            {
                CT_BorderPr pr = _border.IsSetBottom() ? _border.bottom : _border.AddNewBottom();
                if (value == BorderStyle.None) _border.UnsetBottom();
                else pr.style = (ST_BorderStyle)value;
            }
        }

        public BorderStyle BorderDiagonal
        {
            get
            {
                return GetBorderStyle(_border.diagonal);
            }
            set
            {
                CT_BorderPr pr = _border.IsSetDiagonal() ? _border.diagonal : _border.AddNewDiagonal();
                if (value == (short)BorderStyle.None) _border.UnsetDiagonal();
                else pr.style = (ST_BorderStyle)value;
            }
        }

        public BorderStyle BorderLeft
        {
            get
            {
                return GetBorderStyle(_border.left);
            }
            set
            {
                CT_BorderPr pr = _border.IsSetLeft() ? _border.left : _border.AddNewLeft();
                if (value == (short)BorderStyle.None) _border.UnsetLeft();
                else pr.style = (ST_BorderStyle)(value);
            }
        }

        public BorderStyle BorderRight
        {
            get
            {
                return GetBorderStyle(_border.right);
            }
            set
            {
                CT_BorderPr pr = _border.IsSetRight() ? _border.right : _border.AddNewRight();
                if (value == (short)BorderStyle.None) _border.UnsetRight();
                else pr.style = (ST_BorderStyle)(value );
            }
        }

        public BorderStyle BorderTop
        {
            get
            {
                return GetBorderStyle(_border.top);
            }
            set
            {
                CT_BorderPr pr = _border.IsSetTop() ? _border.top : _border.AddNewTop();
                if (value == (short)BorderStyle.None) _border.UnsetTop();
                else pr.style = (ST_BorderStyle)(value );
            }
        }

        public short BottomBorderColor
        {
            get
            {
                return GetIndexedColor(BottomBorderColorColor as XSSFColor);
            }
            set
            {
                CT_Color ctColor = new CT_Color();
                ctColor.indexed = (uint)(value);
                ctColor.indexedSpecified = true;
                SetBottomBorderColor(ctColor);
            }
        }

        public short DiagonalBorderColor
        {
            get
            {
                return GetIndexedColor(DiagonalBorderColorColor as XSSFColor);
            }
            set
            {
                CT_Color ctColor = new CT_Color();
                ctColor.indexed = (uint)(value);
                ctColor.indexedSpecified = true;
                SetDiagonalBorderColor(ctColor);
            }
        }

        public short LeftBorderColor
        {
            get
            {
                return GetIndexedColor(LeftBorderColorColor as XSSFColor);
            }
            set
            {
                CT_Color ctColor = new CT_Color();
                ctColor.indexed = (uint)(value);
                ctColor.indexedSpecified = true;
                SetLeftBorderColor(ctColor);
            }
        }

        public short RightBorderColor
        {
            get
            {
                return GetIndexedColor(RightBorderColorColor as XSSFColor);
            }
            set
            {
                CT_Color ctColor = new CT_Color();
                ctColor.indexed = (uint)(value);
                ctColor.indexedSpecified = true;
                SetRightBorderColor(ctColor);
            }
        }

        public short TopBorderColor
        {
            get
            {
                return GetIndexedColor(RightBorderColorColor as XSSFColor);
            }
            set
            {
                CT_Color ctColor = new CT_Color();
                ctColor.indexed = (uint)(value);
                ctColor.indexedSpecified = true;
                SetTopBorderColor(ctColor);
            }
        }

        public IColor BottomBorderColorColor
        {
            get
            {
                return GetColor(_border.bottom);
            }
            set
            {
                XSSFColor xcolor = XSSFColor.ToXSSFColor(value);
                if (xcolor == null) SetBottomBorderColor((CT_Color)null);
                else SetBottomBorderColor(xcolor.GetCTColor());
            }
        }
        public void SetBottomBorderColor(CT_Color color)
        {
            CT_BorderPr pr = _border.IsSetBottom() ? _border.bottom : _border.AddNewBottom();
            if (color == null)
            {
                pr.UnsetColor();
            }
            else
            {
                pr.color = (color);
            }
        }
        public IColor DiagonalBorderColorColor
        {
            get
            {
                return GetColor(_border.diagonal);
            }
            set
            {
                XSSFColor xcolor = XSSFColor.ToXSSFColor(value);
                if (xcolor == null) SetDiagonalBorderColor((CT_Color)null);
                else SetDiagonalBorderColor(xcolor.GetCTColor());
            }
        }
        public void SetDiagonalBorderColor(CT_Color color)
        {
            CT_BorderPr pr = _border.IsSetDiagonal() ? _border.diagonal : _border.AddNewDiagonal();
            if (color == null)
            {
                pr.UnsetColor();
            }
            else
            {
                pr.color = (color);
            }
        }
        public IColor LeftBorderColorColor
        {
            get
            {
                return GetColor(_border.left);
            }
            set
            {
                XSSFColor xcolor = XSSFColor.ToXSSFColor(value);
                if (xcolor == null) SetLeftBorderColor((CT_Color)null);
                else SetLeftBorderColor(xcolor.GetCTColor());
            }
        }

        public void SetLeftBorderColor(CT_Color color)
        {
            CT_BorderPr pr = _border.IsSetLeft() ? _border.left : _border.AddNewLeft();
            if (color == null)
            {
                pr.UnsetColor();
            }
            else
            {
                pr.color = (color);
            }
        }
        public IColor RightBorderColorColor
        {
            get
            {
                return GetColor(_border.right);
            }
            set
            {
                XSSFColor xcolor = XSSFColor.ToXSSFColor(value);
                if (xcolor == null) SetRightBorderColor((CT_Color)null);
                else SetRightBorderColor(xcolor.GetCTColor());
            }
        }

        public void SetRightBorderColor(CT_Color color)
        {
            CT_BorderPr pr = _border.IsSetRight() ? _border.right : _border.AddNewRight();
            if (color == null)
            {
                pr.UnsetColor();
            }
            else
            {
                pr.color = (color);
            }
        }

        public IColor TopBorderColorColor
        {
            get
            {
                return GetColor(_border.top);
            }
            set
            {
                XSSFColor xcolor = XSSFColor.ToXSSFColor(value);
                if (xcolor == null) SetTopBorderColor((CT_Color)null);
                else SetTopBorderColor(xcolor.GetCTColor());
            }
        }
        public void SetTopBorderColor(CT_Color color)
        {
            CT_BorderPr pr = _border.IsSetTop() ? _border.top : _border.AddNewTop();
            if (color == null)
            {
                pr.UnsetColor();
            }
            else
            {
                pr.color = (color);
            }
        }

        public BorderStyle BorderVertical
        {
            get => GetBorderStyle(_border.vertical);
            set
            {
                CT_BorderPr pr = _border.IsSetVertical() ? _border.vertical : _border.AddNewVertical();
                if (value == BorderStyle.None) _border.UnsetVertical();
                else pr.style = Enums.Parse<ST_BorderStyle>(value.ToString(), true);
            }
        }

        public BorderStyle BorderHorizontal
        {
            get => GetBorderStyle(_border.horizontal);
            set
            {
                CT_BorderPr pr = _border.IsSetHorizontal() ? _border.horizontal : _border.AddNewHorizontal();
                if (value == BorderStyle.None) _border.UnsetHorizontal();
                else pr.style = Enums.Parse<ST_BorderStyle>(value.ToString(), true);
            }
        }


        public short VerticalBorderColor
        {
            get
            {
                return GetIndexedColor(VerticalBorderColorColor as XSSFColor);
            }
            set
            {
                CT_Color ctColor = new CT_Color();
                ctColor.indexed = (uint)value;
                SetVerticalBorderColor(ctColor);
            }
        }

        public IColor VerticalBorderColorColor
        {
            get
            {
                return GetColor(_border.vertical);
            }
            set
            {
                XSSFColor xcolor = XSSFColor.ToXSSFColor(value);
                if (xcolor == null) SetBottomBorderColor((CT_Color)null);
                else SetVerticalBorderColor(xcolor.GetCTColor());
            }
        }
        public void SetVerticalBorderColor(CT_Color color)
        {
            CT_BorderPr pr = _border.IsSetVertical() ? _border.vertical : _border.AddNewVertical();
            if (color == null)
            {
                pr.UnsetColor();
            }
            else
            {
                pr.color = color;
            }
        }

        public short HorizontalBorderColor
        {
            get
            {
                return GetIndexedColor(HorizontalBorderColorColor as XSSFColor);
            }
            set
            {
                CT_Color ctColor = new CT_Color();
                ctColor.indexed = (uint)value;
                SetHorizontalBorderColor(ctColor);
            }
        }

        public IColor HorizontalBorderColorColor
        {
            get
            {
                return GetColor(_border.horizontal);
            }
            set
            {
                XSSFColor xcolor = XSSFColor.ToXSSFColor(value);
                if (xcolor == null) SetBottomBorderColor((CT_Color)null);
                else SetHorizontalBorderColor(xcolor.GetCTColor());
            }
        }


        public void SetHorizontalBorderColor(CT_Color color)
        {
            CT_BorderPr pr = _border.IsSetHorizontal() ? _border.horizontal : _border.AddNewHorizontal();
            if (color == null)
            {
                pr.UnsetColor();
            }
            else
            {
                pr.color = color;
            }
        }



        /**
         * @param borderPr
         * @return BorderStyle from the given element's style, or NONE if border is null
         */
        private static BorderStyle GetBorderStyle(CT_BorderPr borderPr)
        {
            if (borderPr == null) return BorderStyle.None;
            ST_BorderStyle? ptrn = borderPr.style;
            return ptrn == null ? BorderStyle.None : Enums.Parse<BorderStyle>(ptrn.Value.ToString(), true);
        }

        private static short GetIndexedColor(XSSFColor color)
        {
            return (short)(color == null ? 0 : color.Indexed);
        }

        private XSSFColor GetColor(CT_BorderPr pr)
        {
            return pr == null ? null : new XSSFColor(pr.color, _colorMap);
        }
        #endregion
    }
}
