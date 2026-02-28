using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Drawing
{
    /// <summary>
    /// the type of fill to display with the shape or the background of the slide.
    /// </summary>
    public enum MSOFillType
    {
        /// <summary>
        /// A solid fill
        /// </summary>
        Solid = 0,
        /// <summary>
        /// A patterned fill
        /// </summary>
        Pattern = 1,
        /// <summary>
        /// A textured fill
        /// </summary>
        Texture=2,
        /// <summary>
        /// A picture fill
        /// </summary>
        Picture=3,
        /// <summary>
        /// A gradient fill that starts and ends with defined endpoints
        /// </summary>
        Shade=4,
        /// <summary>
        /// A gradient fill that starts and ends based on the bounds of the shape
        /// </summary>
        ShadeCenter=5,
        /// <summary>
        /// A gradient fill that starts on the outline of the shape and ends at a point defined within the shape
        /// </summary>
        ShadeShape=6,
        /// <summary>
        /// A gradient fill that starts on the outline of the shape and ends at a point defined within the shape.
        /// The fill angle is scaled by the aspect ratio of the shape
        /// </summary>
        ShadeScale=7,
        /// <summary>
        /// A gradient fill interpreted by the host application
        /// </summary>
        ShadeTitle=8,
        /// <summary>
        /// A fill that matches the background fill
        /// </summary>
        Background=9
    }
}
