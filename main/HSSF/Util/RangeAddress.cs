/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


using System;
using System.Text;
using System.Globalization;

namespace NPOI.HSSF.Util
{


    /**
     * Title:        Range Address 
     * Description:  provides connectivity utilities for ranges
     *
     *
     * REFERENCE:  
     * @author IgOr KaTz &amp; EuGeNe BuMaGiN (Tal Moshaiov) (VistaPortal LDT.)
    @version 1.0
     */

    public class RangeAddress
    {
        const int WRONG_POS = -1;
        const int MAX_HEIGHT = 66666;
        //static char SO_FORMNAME_ENCLOSURE =  '\'';
        String m_sheetName;
        String m_cellFrom;
        String m_cellTo;

        /**
         * Accepts an external reference from excel.
         * 
         * i.e. Sheet1!$A$4:$B$9
         * @param _url
         */
        public RangeAddress(String _url)
        {
            init(_url);
        }

        public RangeAddress(int _startCol, int _startRow, int _endCol, int _endRow)
        {
            init(NumTo26Sys(_startCol) + _startRow + ":"
            + NumTo26Sys(_endCol) + _endRow);
        }

        /**
         * 
         * @return String <b>note: </b> All absolute references are Removed
         */
        public String Address
        {
            get
            {
                String result = "";
                if (m_sheetName != null)
                    result += m_sheetName + "!";
                if (m_cellFrom != null)
                {
                    result += m_cellFrom;
                    if (m_cellTo != null)
                        result += ":" + m_cellTo;
                }
                return result;
            }
        }


        public String SheetName
        {
            get
            {
                return m_sheetName;
            }
        }

        public String Range
        {
            get
            {
                String result = "";
                if (m_cellFrom != null)
                {
                    result += m_cellFrom;
                    if (m_cellTo != null)
                        result += ":" + m_cellTo;
                }
                return result;
            }
        }

        public bool IsCellOk(String _cell)
        {
            if (_cell != null)
            {
                if ((GetYPosition(_cell) != WRONG_POS) &&
                (GetXPosition(_cell) != WRONG_POS))
                    return true;
                else
                    return false;
            }
            else
                return false;
        }

        public bool IsSheetNameOk()
        {

            return IsSheetNameOk(m_sheetName);
        }

        private static bool intern_isSheetNameOk(String _sheetName, bool _canBeWaitSpace)
        {
            for (int i = 0; i < _sheetName.Length; i++)
            {
                char ch = _sheetName[i];
                if (!(Char.IsLetterOrDigit(ch) || (ch == '_') ||
                _canBeWaitSpace && (ch == ' ')))
                {
                    return false;
                }
            }
            return true;
        }

        public static bool IsSheetNameOk(String _sheetName)
        {
            bool res = false;
            if (!string.IsNullOrEmpty(_sheetName))
            {
                res = intern_isSheetNameOk(_sheetName, true);
            }
            else
                res = true;
            return res;
        }


        public String FromCell
        {
            get
            {
                return m_cellFrom;
            }
        }

        public String ToCell
        {
            get
            {
                return m_cellTo;
            }
        }

        public int Width
        {
            get
            {
                if (m_cellFrom != null && m_cellTo != null)
                {
                    int toX = GetXPosition(m_cellTo);
                    int fromX = GetXPosition(m_cellFrom);
                    if ((toX == WRONG_POS) || (fromX == WRONG_POS))
                    {
                        return 0;
                    }
                    else
                        return toX - fromX + 1;
                }
                return 0;
            }
        }

        public int Height
        {
            get
            {
                if (m_cellFrom != null && m_cellTo != null)
                {
                    int toY = GetYPosition(m_cellTo);
                    int fromY = GetYPosition(m_cellFrom);
                    if ((toY == WRONG_POS) || (fromY == WRONG_POS))
                    {
                        return 0;
                    }
                    else
                        return toY - fromY + 1;
                }
                return 0;
            }
        }

        public void SetSize(int _width, int _height)
        {
            if (m_cellFrom == null)
                m_cellFrom = "a1";
            int tlX, tlY; // fix warning CS0168 "never used": , rbX, rbY; //Tony Qu: not sure what's doing here
            tlX = GetXPosition(m_cellFrom);
            tlY = GetYPosition(m_cellFrom);
            m_cellTo = NumTo26Sys(tlX + _width - 1);
            m_cellTo += (tlY + _height - 1).ToString(CultureInfo.InvariantCulture);
        }

        public bool HasSheetName
        {
            get
            {
                if (m_sheetName == null)
                    return false;
                return true;
            }
        }

        public bool HasRange
        {
            get
            {
                return (m_cellFrom != null && m_cellTo != null && !m_cellFrom.Equals(m_cellTo));
            }
        }

        public bool HasCell
        {
            get
            {
                if (m_cellFrom == null)
                    return false;
                return true;
            }
        }

        private void init(String _url)
        {

            _url = RemoveString(_url, "$");
            _url = RemoveString(_url, "'");

            String[] urls = ParseURL(_url);
            m_sheetName = urls[0];
            m_cellFrom = urls[1];
            m_cellTo = urls[2];

            //What if range is one celled ?
            if (m_cellTo == null)
            {
                m_cellTo = m_cellFrom;
            }

            //Removing noneeds Chars
            m_cellTo = RemoveString(m_cellTo, ".");


        }

        private String[] ParseURL(String _url)
        {
            String[] result = new String[3];
            int index = _url.IndexOf(':');
            if (index >= 0)
            {
                String fromStr = _url.Substring(0, index);
                String toStr = _url.Substring(index + 1);
                index = fromStr.IndexOf('!');
                if (index >= 0)
                {
                    result[0] = fromStr.Substring(0, index);
                    result[1] = fromStr.Substring(index + 1);
                }
                else
                {
                    result[1] = fromStr;
                }
                index = toStr.IndexOf('!');
                if (index >= 0)
                {
                    result[2] = toStr.Substring(index + 1);
                }
                else
                {
                    result[2] = toStr;
                }
            }
            else
            {
                index = _url.IndexOf('!');
                if (index >= 0)
                {
                    result[0] = _url.Substring(0, index);
                    result[1] = _url.Substring(index + 1);
                }
                else
                {
                    result[1] = _url;
                }
            }
            return result;
        }

        public int GetYPosition(String _subrange)
        {
            int result = WRONG_POS;
            _subrange = _subrange.Trim();
            if (_subrange.Length != 0)
            {
                String digitstr = GetDigitPart(_subrange);
                try
                {
                    result = int.Parse(digitstr, CultureInfo.InvariantCulture);
                    if (result > MAX_HEIGHT)
                    {
                        result = WRONG_POS;
                    }
                }
                catch (Exception)
                {

                    result = WRONG_POS;
                }
            }
            return result;
        }

        private static bool IsLetter(String _str)
        {
            bool res = true;
            if (!string.IsNullOrEmpty(_str))
            {
                for (int i = 0; i < _str.Length; i++)
                {
                    char ch = _str[i];
                    if (!Char.IsLetter(ch))
                    {
                        res = false;
                        break;
                    }
                }
            }
            else
                res = false;
            return res;
        }

        public int GetXPosition(String _subrange)
        {
            int result = WRONG_POS;
            String tmp = Filter(_subrange);
            tmp = this.GetCharPart(_subrange);
            // we will Process only 2 letters ranges
            if (IsLetter(tmp) && ((tmp.Length == 2) || (tmp.Length == 1)))
            {
                result = Get26Sys(tmp);
            }
            return result;
        }

        public String GetDigitPart(String _value)
        {
            String result = "";
            int digitpos = GetFirstDigitPosition(_value);
            if (digitpos >= 0)
            {
                result = _value.Substring(digitpos);
            }
            return result;
        }

        public String GetCharPart(String _value)
        {
            String result = "";
            int digitpos = GetFirstDigitPosition(_value);
            if (digitpos >= 0)
            {
                result = _value.Substring(0, digitpos);
            }
            return result;
        }

        private String Filter(String _range)
        {
            String res = "";
            for (int i = 0; i < _range.Length; i++)
            {
                char ch = _range[i];
                if (ch != '$')
                {
                    res = res + ch;
                }
            }
            return res;
        }

        private int GetFirstDigitPosition(String _value)
        {
            int result = WRONG_POS;
            if (_value != null && _value.Trim().Length == 0)
            {
                return result;
            }
            _value = _value.Trim();
            int Length = _value.Length;
            for (int i = 0; i < Length; i++)
            {
                if (Char.IsDigit(_value[i]))
                {
                    result = i;
                    break;
                }
            }
            return result;
        }

        public int Get26Sys(String _s)
        {
            int sum = 0;
            int multiplier = 1;
            if (!string.IsNullOrEmpty(_s))
            {
                for (int i = _s.Length - 1; i >= 0; i--)
                {
                    char ch = _s[i];
                    int val = (int)(Char.GetNumericValue(ch) - Char.GetNumericValue('A') + 1);
                    sum = sum + val * multiplier;
                    multiplier = multiplier * 26;
                }
                return sum;
            }
            return WRONG_POS;
        }

        public String NumTo26Sys(int _num)
        {
            //int sum = 0;
            int reminder;
            String s = "";
            do
            {
                _num--;
                reminder = _num % 26;
                int val = 65 + reminder;
                _num = _num / 26;
                s = (char)val + s; // reverce
            } while (_num > 0);
            return s;
        }

        public String ReplaceString(String _source, String _oldPattern,
        String _newPattern)
        {
            StringBuilder res = new StringBuilder(_source);
            res = res.Replace(_oldPattern, _newPattern);

            return res.ToString();
        }

        public String RemoveString(String _source, String _match)
        {
            return ReplaceString(_source, _match, "");
        }

    }
}