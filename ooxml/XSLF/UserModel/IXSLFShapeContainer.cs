using System.Collections.Generic;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A PresentationML sheet or group containing ordered child shapes.
    /// </summary>
    public interface IXSLFShapeContainer
    {
        /// <summary>
        /// Returns a read-only snapshot of immediate children in XML drawing order.
        /// Unsupported drawing elements are included as opaque wrappers.
        /// </summary>
        IReadOnlyList<XSLFShape> GetShapes();

        /// <summary>
        /// Appends a preset shape with a sheet-wide unique identifier and an empty anchor.
        /// </summary>
        XSLFSimpleShape CreateAutoShape(ST_ShapeType shapeType = ST_ShapeType.rect);

        /// <summary>
        /// Appends an empty group with a sheet-wide unique identifier.
        /// Exterior and interior anchors are initially empty and can be set independently.
        /// </summary>
        XSLFGroupShape CreateGroup();
    }
}
