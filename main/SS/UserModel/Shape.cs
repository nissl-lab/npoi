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

namespace NPOI.SS.UserModel
{
    public interface IShape
    {
        string ShapeName { get; }
        IChildAnchor Anchor { get; }

        IShape Parent { get; }

        uint ID { get; }
        string Name { get; set; }
        void SetLineStyleColor(int red, int green, int blue);
        void SetFillColor(int red, int green, int blue);

        int LineStyleColor { get; }

        int FillColor { get; set; }
        double LineWidth { get; set; }
        LineStyle LineStyle { get; set; }
        LineEndingCapType LineEndingCapType{ get; set; }
        CompoundLineType CompoundLineType { get; set; }
        bool IsNoFill { get; set; }
        int CountOfAllChildren { get; }
    }
}
