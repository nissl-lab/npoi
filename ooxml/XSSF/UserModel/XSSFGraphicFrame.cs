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

using NPOI.OpenXmlFormats.Dml;
using System;
using NPOI.OpenXmlFormats;
namespace NPOI.XSSF.UserModel
{

    /**
     * Represents DrawingML GraphicalObjectFrame.
     *
     * @author Roman Kashitsyn
     */
    public class XSSFGraphicFrame
    {

        private static CT_GraphicalObjectFrame prototype = null;

        private CT_GraphicalObjectFrame graphicFrame;
        private XSSFDrawing Drawing;
        private XSSFClientAnchor anchor;

        /**
         * Construct a new XSSFGraphicFrame object.
         *
         * @param Drawing the XSSFDrawing that owns this frame
         * @param ctGraphicFrame the XML bean that stores this frame content
         */
        protected XSSFGraphicFrame(XSSFDrawing Drawing, CT_GraphicalObjectFrame ctGraphicFrame)
        {
            this.Drawing = Drawing;
            this.graphicFrame = ctGraphicFrame;
        }


        public CT_GraphicalObjectFrame GetCTGraphicalObjectFrame()
        {
            return graphicFrame;
        }

        /**
         * Initialize default structure of a new graphic frame
         */
        public static CT_GraphicalObjectFrame Prototype()
        {
            if (prototype == null)
            {
                CT_GraphicalObjectFrame graphicFrame = CT_GraphicalObjectFrame.Factory.newInstance();

                CT_GraphicalObjectFrameNonVisual nvGraphic = graphicFrame.AddNewNvGraphicFramePr();
                CT_NonVisualDrawingProps props = nvGraphic.AddNewCNvPr();
                props.id = (0);
                props.name = ("Diagramm 1");
                nvGraphic.AddNewCNvGraphicFramePr();

                CT_Transform2D transform = graphicFrame.AddNewXfrm();
                CT_PositiveSize2D extPoint = transform.AddNewExt();
                CT_Point2D offPoint = transform.AddNewOff();

                extPoint.cx=(0);
                extPoint.cy=(0);
                offPoint.x=(0);
                offPoint.y=(0);

                CT_GraphicalObject graphic = graphicFrame.AddNewGraphic();

                prototype = graphicFrame;
            }
            return prototype;
        }

        /**
         * Sets the frame macro.
         */
        public void SetMacro(String macro)
        {
            graphicFrame.macro = (macro);
        }

        /**
         * Sets the frame name.
         */
        public void SetName(String name)
        {
            GetNonVisualProperties().name = (name);
        }

        /**
         * Returns the frame name.
         * @return name of the frame
         */
        public String GetName()
        {
            return GetNonVisualProperties().name;
        }

        private CT_NonVisualDrawingProps GetNonVisualProperties()
        {
            CT_GraphicalObjectFrameNonVisual nvGraphic = graphicFrame.nvGraphicFramePr;
            return nvGraphic.cNvPr;
        }

        /**
         * Attaches frame to an anchor.
         */
        public void SetAnchor(XSSFClientAnchor anchor)
        {
            this.anchor = anchor;
        }

        /**
         * Returns the frame anchor.
         * @return the anchor this frame is attached to
         */
        public XSSFClientAnchor GetAnchor()
        {
            return anchor;
        }

        /**
         * Assign a DrawingML chart to the graphic frame.
         */
        //protected void SetChart(XSSFChart chart, String relId)
        //{
        //    CT_GraphicalObjectData data = graphicFrame.graphic.AddNewGraphicData();
        //    AppendChartElement(data, relId);
        //    chart.SetGraphicFrame(this);
        //    return;
        //}

        /**
         * Gets the frame id.
         */
        public long GetId()
        {
            return graphicFrame.nvGraphicFramePr.cNvPr.id;
        }

        /**
         * Sets the frame id.
         */
        public void SetId(long id)
        {
            graphicFrame.nvGraphicFramePr.cNvPr.id = (uint)id;
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
        private void AppendChartElement(CT_GraphicalObjectData data, String id)
        {
            String r_namespaceUri = ST_RelationshipId.NamespaceURI;
            String c_namespaceUri = XSSFDrawing.NAMESPACE_C;
            XmlCursor cursor = data.newCursor();
            cursor.ToNextToken();
            cursor.beginElement(new QName(c_namespaceUri, "chart", "c"));
            cursor.insertAttributeWithValue(new QName(r_namespaceUri, "id", "r"), id);
            cursor.dispose();
            data.SetUri(c_namespaceUri);
        }

    }
}



