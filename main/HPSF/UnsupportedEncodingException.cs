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
using System.IO;

namespace NPOI.HPSF
{
    /**
     * The Character Encoding is not supported.
     *
     * @author  Asmus Freytag
     * @since   JDK1.1
     */
    public class UnsupportedEncodingException : IOException
    {
        //private static long serialVersionUID = -4274276298326136670L;

        /**
         * Constructs an UnsupportedEncodingException without a detail message.
         */
        public UnsupportedEncodingException()
            : base()
        {

        }

        /**
         * Constructs an UnsupportedEncodingException with a detail message.
         * @param s Describes the reason for the exception.
         */
        public UnsupportedEncodingException(String s)
            : base(s)
        {

        }
    }
}