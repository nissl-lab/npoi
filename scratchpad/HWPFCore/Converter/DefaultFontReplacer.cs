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

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Converter
{
    public class DefaultFontReplacer : FontReplacer
    {
        #region IFontReplacer 成员

        public Triplet Update(Triplet original)
        {
            if (!string.IsNullOrEmpty(original.fontName))
            {
                String fontName = original.fontName;

                if (fontName.EndsWith(" Regular"))
                    fontName = AbstractWordUtils.SubstringBeforeLast(fontName,
                            " Regular");

                if (fontName
                        .EndsWith(" \u041F\u043E\u043B\u0443\u0436\u0438\u0440\u043D\u044B\u0439"))
                    fontName = AbstractWordUtils
                            .SubstringBeforeLast(fontName,
                                    " \u041F\u043E\u043B\u0443\u0436\u0438\u0440\u043D\u044B\u0439")
                            + " Bold";

                if (fontName
                        .EndsWith(" \u041F\u043E\u043B\u0443\u0436\u0438\u0440\u043D\u044B\u0439 \u041A\u0443\u0440\u0441\u0438\u0432"))
                    fontName = AbstractWordUtils
                            .SubstringBeforeLast(
                                    fontName,
                                    " \u041F\u043E\u043B\u0443\u0436\u0438\u0440\u043D\u044B\u0439 \u041A\u0443\u0440\u0441\u0438\u0432")
                            + " Bold Italic";

                if (fontName.EndsWith(" \u041A\u0443\u0440\u0441\u0438\u0432"))
                    fontName = AbstractWordUtils.SubstringBeforeLast(fontName,
                            " \u041A\u0443\u0440\u0441\u0438\u0432") + " Italic";

                original.fontName = fontName;
            }

            if (!string.IsNullOrEmpty(original.fontName))
            {
                if ("Times Regular".Equals(original.fontName)
                        || "Times-Regular".Equals(original.fontName))
                {
                    original.fontName = "Times";
                    original.bold = false;
                    original.italic = false;
                }
                if ("Times Bold".Equals(original.fontName)
                        || "Times-Bold".Equals(original.fontName))
                {
                    original.fontName = "Times";
                    original.bold = true;
                    original.italic = false;
                }
                if ("Times Italic".Equals(original.fontName)
                        || "Times-Italic".Equals(original.fontName))
                {
                    original.fontName = "Times";
                    original.bold = false;
                    original.italic = true;
                }
                if ("Times Bold Italic".Equals(original.fontName)
                        || "Times-BoldItalic".Equals(original.fontName))
                {
                    original.fontName = "Times";
                    original.bold = true;
                    original.italic = true;
                }
            }

            return original;
        }

        #endregion
    }
}
