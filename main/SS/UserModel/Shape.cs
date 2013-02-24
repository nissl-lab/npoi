using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public interface IShape
    {
        IShape Parent { get; }

        void SetLineStyleColor(int lineStyleColor);
        void SetLineStyleColor(int red, int green, int blue);
        void SetFillColor(int red, int green, int blue);

        int LineStyleColor { get; }

        int FillColor { get; set; }
        double LineWidth { get; set; }
        LineStyle LineStyle { get; set; }
        bool IsNoFill { get; set; }
        int CountOfAllChildren { get; }
    }
}
