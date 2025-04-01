/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


namespace NPOI.XDDF.UserModel.Text
{
    using NPOI.Common.UserModel.Fonts;
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFFont
    {
        private FontGroup group;
        private CT_TextFont font;

        public static XDDFFont UnsetFontForGroup(FontGroup group)
        {
            return new XDDFFont(group, null);
        }

        public XDDFFont(FontGroup group, string typeface, sbyte? charset, sbyte? pitch, byte[] panose)
                : this(group, new CT_TextFont())
        {

            font.typeface = typeface;
            if(charset == null)
            {
                font.charsetSpecified = false;
            }
            else
            {
                font.charset = charset.Value;
            }
            if(pitch == null)
            {
                font.pitchFamilySpecified = false;
            }
            else
            {
                font.pitchFamily = pitch.Value;
            }
            if(panose == null || panose.Length == 0)
            {
                font.panoseSpecified = false;
            }
            else
            {
                font.panose = panose;
            }
        }
        internal XDDFFont(FontGroup group, CT_TextFont font)
        {
            this.group = group;
            this.font = font;
        }
        internal CT_TextFont GetXmlObject()
        {
            return font;
        }

        public FontGroup Group
        {
            get
            {
                return group;
            }
        }

        public string Typeface
        {
            get
            {
                return font.typeface;
            }
        }

        public sbyte? Charset
        {
            get
            {
                if(font.charsetSpecified)
                {
                    return font.charset;
                }
                else
                {
                    return null;
                }
            }
        }

        public sbyte? PitchFamily
        {
            get
            {
                if(font.pitchFamilySpecified)
                {
                    return font.pitchFamily;
                }
                else
                {
                    return null;
                }
            }
        }

        public byte[] Panose
        {
            get
            {
                if(font.panoseSpecified)
                {
                    return font.panose;
                }
                else
                {
                    return null;
                }
            }
        }
    }
}


