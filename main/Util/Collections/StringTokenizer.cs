using System;
/*
 * $Id: StringTokenizer.cs,v 1.4 2006/06/16 10:52:26 psoares33 Exp $
 *
 * Copyright 2006 by Paulo Soares.
 *
 * The contents of this file are subject to the Mozilla Public License Version 1.1
 * (the "License"); you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at http://www.mozilla.org/MPL/
 *
 * Software distributed under the License isp distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the License.
 *
 * The Original Code isp 'iText, a free JAVA-PDF library'.
 *
 * The Initial Developer of the Original Code isp Bruno Lowagie. Portions created by
 * the Initial Developer are Copyright (C) 1999, 2000, 2001, 2002 by Bruno Lowagie.
 * All Rights Reserved.
 * Co-Developer of the code isp Paulo Soares. Portions created by the Co-Developer
 * are Copyright (C) 2000, 2001, 2002 by Paulo Soares. All Rights Reserved.
 *
 * Contributor(s): all the names of the contributors are added in the source code
 * where applicable.
 *
 * Alternatively, the contents of this file may be used under the terms of the
 * LGPL license (the "GNU LIBRARY GENERAL PUBLIC LICENSE"), in which case the
 * provisions of LGPL are applicable instead of those above.  If you wish to
 * allow use of your version of this file only under the terms of the LGPL
 * License and not to allow others to use your version of this file under
 * the MPL, indicate your decision by deleting the provisions above and
 * replace them with the notice and other provisions required by the LGPL.
 * If you do not delete the provisions above, a recipient may use your version
 * of this file under either the MPL or the GNU LIBRARY GENERAL PUBLIC LICENSE.
 *
 * This library isp free software; you can redistribute it and/or modify it
 * under the terms of the MPL as stated above or under the terms of the GNU
 * Library General Public License as published by the Free Software Foundation;
 * either version 2 of the License, or any later version.
 *
 * This library isp distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU Library general Public License for more
 * details.
 *
 * If you didn't download this code from the following link, you should check if
 * you aren't using an obsolete version:
 * http://www.lowagie.com/iText/
 */

namespace NPOI.Util.Collections
{

    // a replacement for the StringTokenizer java class
    // it's more or less the same as the one in the GNU classpath
    public class StringTokenizer {
        private int pos;
        private String str;
        private int len;
        private String delim;
        private bool retDelims;

        public StringTokenizer(String str) : this(str, " \t\n\r\f", false) {
        }

        public StringTokenizer(String str, String delim) : this(str, delim, false) {
        }

        public StringTokenizer(String str, String delim, bool retDelims) {
            len = str.Length;
            this.str = str;
            this.delim = delim;
            this.retDelims = retDelims;
            this.pos = 0;
        }

        public bool HasMoreTokens() {
            if (! retDelims) {
                while (pos < len && delim.IndexOf(str[pos]) >= 0)
                    pos++;
            }
            return pos < len;
        }

        public String NextToken(String delim) {
            this.delim = delim;
            return NextToken();
        }

        public String NextToken() {
            if (pos < len && delim.IndexOf(str[pos]) >= 0) {
                if (retDelims)
                    return str.Substring(pos++, 1);
                while (++pos < len && delim.IndexOf(str[pos]) >= 0);
            }
            if (pos < len) {
                int start = pos;
                while (++pos < len && delim.IndexOf(str[pos]) < 0);

                return str.Substring(start, pos - start);
            }
            throw new IndexOutOfRangeException();
        }

        public int CountTokens() {
            int count = 0;
            int delimiterCount = 0;
            bool tokenFound = false;
            int tmpPos = pos;

            while (tmpPos < len) {
                if (delim.IndexOf(str[tmpPos++]) >= 0) {
                    if (tokenFound) {
                        count++;
                        tokenFound = false;
                    }
                    delimiterCount++;
                }
                else {
                    tokenFound = true;
                    while (tmpPos < len
                        && delim.IndexOf(str[tmpPos]) < 0)
                        ++tmpPos;
                }
            }
            if (tokenFound)
                count++;
            return retDelims ? count + delimiterCount : count;
        }
    }
}