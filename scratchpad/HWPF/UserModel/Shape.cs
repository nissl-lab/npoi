/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.UserModel
{
    using System;
    using NPOI.HWPF.Model;
    using NPOI.Util;

    public class Shape
    {
        int _id, _left, _right, _top, _bottom;
        /**
         * true if the Shape bounds are within document (for
         * example, it's false if the image left corner Is outside the doc, like for
         * embedded documents)
         */
        bool _inDoc;

        public Shape(GenericPropertyNode nodo)
        {
            byte[] contenuto = nodo.Bytes;
            _id = LittleEndian.GetInt(contenuto);
            _left = LittleEndian.GetInt(contenuto, 4);
            _top = LittleEndian.GetInt(contenuto, 8);
            _right = LittleEndian.GetInt(contenuto, 12);
            _bottom = LittleEndian.GetInt(contenuto, 16);
            _inDoc = (_left >= 0 && _right >= 0 && _top >= 0 && _bottom >=
0);
        }

        public int Id
        {
            get
            {
                return _id;
            }
        }

        public int Left
        {
            get
            {
                return _left;
            }
        }

        public int Right
        {
            get
            {
                return _right;
            }
        }

        public int Top
        {
            get
            {
                return _top;
            }
        }

        public int Bottom
        {
            get
            {
                return _bottom;
            }
        }

        public int Width
        {
            get
            {
                return _right - _left + 1;
            }
        }

        public int Height
        {
            get
            {
                return _bottom - _top + 1;
            }
        }

        public bool IsWithinDocument
        {
            get
            {
                return _inDoc;
            }
        }
    }
}