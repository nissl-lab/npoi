/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Record
{
	public enum TextOrientation {
		None,
		TopToBottom,
		RotRight,
		RotLeft
	}

    public enum HorizontalTextAlignment:int 
    {
        Left=1,
        Center=2,
        Right=3,
        Justify = 4
    }
    public enum VerticalTextAlignment:int
    {
        Top = 1,
        Center = 2,
        Bottom = 3,
        Justify = 4
    }
}
