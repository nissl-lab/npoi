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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;


namespace TestCases.HSSF.Model
{
    /**
     * @author Evgeniy Berlog
     * @date 25.06.12
     */
    public class HSSFTestModelHelper
    {
        public static TextboxShape CreateTextboxShape(int shapeId, HSSFTextbox textbox)
        {
            return new TextboxShape(textbox, shapeId);
        }

        public static CommentShape CreateCommentShape(int shapeId, HSSFComment comment)
        {
            return new CommentShape(comment, shapeId);
        }

        public static PolygonShape CreatePolygonShape(int shapeId, HSSFPolygon polygon)
        {
            return new PolygonShape(polygon, shapeId);
        }
    }
}
