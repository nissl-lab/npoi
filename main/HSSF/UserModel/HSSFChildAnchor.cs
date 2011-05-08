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


namespace NPOI.HSSF.UserModel
{
    using NPOI.DDF;

    public class HSSFChildAnchor : HSSFAnchor
    {
        public HSSFChildAnchor()
        {
        }

        public HSSFChildAnchor(int dx1, int dy1, int dx2, int dy2):base(dx1, dy1, dx2, dy2)
        {
            
        }

        public void SetAnchor(int dx1, int dy1, int dx2, int dy2)
        {
            this.Dx1 = dx1;
            this.Dy1 = dy1;
            this.Dx2 = dx2;
            this.Dy2 = dy2;
        }

        public override bool IsHorizontallyFlipped
        {
            get { return Dx1 > Dx2; }
        }

        public override bool IsVerticallyFlipped
        {
            get{return Dy1 > Dy2;}
        }

    }
}