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



/*
 * DateUtil.java
 *
 * Created on January 19, 2002, 9:30 AM
 */
namespace NPOI.HSSF.UserModel
{
    using System;



    using NPOI.SS.UserModel;

    /**
     * Contains methods for dealing with Excel dates.
     *
     * @author  Michael Harhen
     * @author  Glen Stampoultzis (glens at apache.org)
     * @author  Dan Sherman (dsherman at isisph.com)
     * @author  Hack Kampbjorn (hak at 2mba.dk)
     * @author  Alex Jacoby (ajacoby at gmail.com)
     * @author  Pavel Krupets (pkrupets at palmtreebusiness dot com)
     */

    public class HSSFDateUtil : DateUtil
    {
        protected new static int AbsoluteDay(DateTime cal, bool use1904windowing)
        {
            return DateUtil.AbsoluteDay(cal, use1904windowing);
        }
    }
}
