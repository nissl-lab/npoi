using System.Collections.Generic;

namespace NPOI.SL.Draw.Geom
{
    /// <summary>
    /// Holds the ordered adjustment and guide formulas of a custom shape geometry.
    /// </summary>
    public class CustomGeometry
    {
        /// <summary>
        /// Gets the adjustment formulas evaluated before the dependent guides.
        /// </summary>
        public IList<GuideIf> Adjusts { get; } = new List<GuideIf>();

        /// <summary>
        /// Gets the guide formulas in evaluation order.
        /// </summary>
        public IList<GuideIf> Guides { get; } = new List<GuideIf>();
    }
}
