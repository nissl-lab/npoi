/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace NPOI.xssf.usermodel;

using javax.xml.namespace.QName;

using NPOI.Openxml4j.opc.PackageRelationship;
using NPOI.util.Internal;
using org.apache.xmlbeans.XmlObject;
using org.apache.xmlbeans.XmlCursor;
using org.Openxmlformats.schemas.Drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame;
using org.Openxmlformats.schemas.Drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrameNonVisual;
using org.Openxmlformats.schemas.Drawingml.x2006.main.CTGraphicalObject;
using org.Openxmlformats.schemas.Drawingml.x2006.main.CTGraphicalObjectData;
using org.Openxmlformats.schemas.Drawingml.x2006.main.CTNonVisualDrawingProps;
using org.Openxmlformats.schemas.Drawingml.x2006.main.CTTransform2D;
using org.Openxmlformats.schemas.Drawingml.x2006.main.CTPoint2D;
using org.Openxmlformats.schemas.Drawingml.x2006.main.CTPositiveSize2D;
using org.Openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;

/**
 * Represents DrawingML GraphicalObjectFrame.
 *
 * @author Roman Kashitsyn
 */
public class XSSFGraphicFrame {

	private static CTGraphicalObjectFrame prototype = null;

	private CTGraphicalObjectFrame graphicFrame;
	private XSSFDrawing Drawing;
	private XSSFClientAnchor anchor;

	/**
	 * Construct a new XSSFGraphicFrame object.
	 *
	 * @param Drawing the XSSFDrawing that owns this frame
	 * @param ctGraphicFrame the XML bean that stores this frame content
	 */
	protected XSSFGraphicFrame(XSSFDrawing Drawing, CTGraphicalObjectFrame ctGraphicFrame) {
		this.Drawing = Drawing;
		this.graphicFrame = ctGraphicFrame;
	}

	
	public CTGraphicalObjectFrame GetCTGraphicalObjectFrame() {
		return graphicFrame;
	}

	/**
	 * Initialize default structure of a new graphic frame
	 */
	protected static CTGraphicalObjectFrame prototype() {
		if (prototype == null) {
			CTGraphicalObjectFrame graphicFrame = CTGraphicalObjectFrame.Factory.newInstance();

			CTGraphicalObjectFrameNonVisual nvGraphic = graphicFrame.AddNewNvGraphicFramePr();
			CTNonVisualDrawingProps props = nvGraphic.AddNewCNvPr();
			props.SetId(0);
			props.SetName("Diagramm 1");
			nvGraphic.AddNewCNvGraphicFramePr();

			CTTransform2D transform = graphicFrame.AddNewXfrm();
			CTPositiveSize2D extPoint = transform.AddNewExt();
			CTPoint2D offPoint = transform.AddNewOff();

			extPoint.SetCx(0);
			extPoint.SetCy(0);
			offPoint.SetX(0);
			offPoint.SetY(0);

			CTGraphicalObject graphic = graphicFrame.AddNewGraphic();

			prototype = graphicFrame;
		}
		return prototype;
	}

	/**
	 * Sets the frame macro.
	 */
	public void SetMacro(String macro) {
		graphicFrame.SetMacro(macro);
	}

	/**
	 * Sets the frame name.
	 */
	public void SetName(String name) {
	 GetNonVisualProperties().SetName(name);
	}

	/**
	 * Returns the frame name.
	 * @return name of the frame
	 */
	public String GetName() {
		return GetNonVisualProperties().GetName();
	}

	private CTNonVisualDrawingProps GetNonVisualProperties() {
		CTGraphicalObjectFrameNonVisual nvGraphic = graphicFrame.GetNvGraphicFramePr();
		return nvGraphic.GetCNvPr();
	}

	/**
	 * Attaches frame to an anchor.
	 */
	protected void SetAnchor(XSSFClientAnchor anchor) {
		this.anchor = anchor;
	}

	/**
	 * Returns the frame anchor.
	 * @return the anchor this frame is attached to
	 */
	public XSSFClientAnchor GetAnchor() {
		return anchor;
	}

	/**
	 * Assign a DrawingML chart to the graphic frame.
	 */
	protected void SetChart(XSSFChart chart, String relId) {
		CTGraphicalObjectData data = graphicFrame.GetGraphic().AddNewGraphicData();
		AppendChartElement(data, relId);
		chart.SetGraphicFrame(this);
		return;
	}

	/**
	 * Gets the frame id.
	 */
	public long GetId() {
		return graphicFrame.GetNvGraphicFramePr().GetCNvPr().GetId();
	}

	/**
	 * Sets the frame id.
	 */
	protected void SetId(long id) {
		graphicFrame.GetNvGraphicFramePr().GetCNvPr().SetId(id);
	}

	/**
	 * The low level code to insert {@code <c:chart>} tag into
	 * {@code<a:graphicData>}.
	 *
	 * Here is the schema (ECMA-376):
	 * <pre>
	 * {@code
	 * <complexType name="CT_GraphicalObjectData">
	 *   <sequence>
	 *     <any minOccurs="0" maxOccurs="unbounded" ProcessContents="strict"/>
	 *   </sequence>
	 *   <attribute name="uri" type="xsd:token"/>
	 * </complexType>
	 * }
	 * </pre>
	 */
	private void AppendChartElement(CTGraphicalObjectData data, String id) {
		String r_namespaceUri = STRelationshipId.type.GetName().GetNamespaceURI();
		String c_namespaceUri = XSSFDrawing.NAMESPACE_C;
		XmlCursor cursor = data.newCursor();
		cursor.ToNextToken();
		cursor.beginElement(new QName(c_namespaceUri, "chart", "c"));
		cursor.insertAttributeWithValue(new QName(r_namespaceUri, "id", "r"), id);
		cursor.dispose();
		data.SetUri(c_namespaceUri);
	}

}



