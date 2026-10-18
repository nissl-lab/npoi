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
        /// Appends a text box with one empty paragraph and a sheet-wide unique identifier.
        /// </summary>
        XSLFTextBox CreateTextBox();

        /// <summary>
        /// Appends an embedded PNG/JPEG picture from this presentation with an empty point-based anchor.
        /// </summary>
        XSLFPictureShape CreatePicture(XSLFPictureData pictureData);

        /// <summary>
        /// Appends a fixed positive row/column grid with zero frame, column and row dimensions.
        /// Frame and grid sizes are set independently in points; no automatic layout is performed.
        /// </summary>
        XSLFTable CreateTable(int rows, int columns);

        /// <summary>
        /// Appends an empty-anchor frame referencing an existing chart in this presentation.
        /// Does not create or modify chart contents, workbooks or external data.
        /// </summary>
        XSLFChartShape CreateChart(XSLFChart chart);

        /// <summary>
        /// Appends an empty group with a sheet-wide unique identifier.
        /// Exterior and interior anchors are initially empty and can be set independently.
        /// </summary>
        XSLFGroupShape CreateGroup();
    }
}
