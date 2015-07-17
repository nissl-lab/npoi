/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using System.Xml;

namespace NPOI.XSSF.UserModel
{
    using NPOI.SS.UserModel;
    using NPOI.OpenXmlFormats.Dml.Chart;
    using System.Collections.Generic;
    using NPOI.OpenXml4Net.OPC;
    using System;
    using NPOI.OpenXmlFormats;
    using System.IO;
    using System.Text;
    using NPOI.SS.UserModel.Charts;
    using NPOI.XSSF.UserModel.Charts;
    using System.Xml.Serialization;
    using NPOI.OpenXmlFormats.Dml;

    /**
     * Represents a SpreadsheetML Chart
     * @author Nick Burch
     * @author Roman Kashitsyn
     */
    public class XSSFChart : POIXMLDocumentPart, IChart, IChartAxisFactory
    {

        /**
         * Parent graphic frame.
         */
        private XSSFGraphicFrame frame;

        /**
         * Root element of the SpreadsheetML Chart part
         */
        private ChartSpaceDocument chartSpaceDocument;
        /**
         * The Chart within that
         */
        private CT_Chart chart;

        List<IChartAxis> axis = new List<IChartAxis>();

        /**
         * Create a new SpreadsheetML chart
         */
        public XSSFChart()
            : base()
        {
            CreateChart();
        }

        /**
         * Construct a SpreadsheetML chart from a namespace part.
         *
         * @param part the namespace part holding the chart data,
         * the content type must be <code>application/vnd.Openxmlformats-officedocument.Drawingml.chart+xml</code>
         * @param rel  the namespace relationship holding this chart,
         * the relationship type must be http://schemas.Openxmlformats.org/officeDocument/2006/relationships/chart
         */
        protected XSSFChart(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            XmlDocument doc = ConvertStreamToXml(part.GetInputStream());
            chartSpaceDocument = ChartSpaceDocument.Parse(doc, NamespaceManager);
            chart = chartSpaceDocument.GetChartSpace().chart;
        }

        /**
         * Construct a new CTChartSpace bean.
         * By default, it's just an empty placeholder for chart objects.
         *
         * @return a new CTChartSpace bean
         */
        private void CreateChart()
        {
            chartSpaceDocument = new ChartSpaceDocument();
            chart = chartSpaceDocument.GetChartSpace().AddNewChart();
            CT_PlotArea plotArea = chart.AddNewPlotArea();

            plotArea.AddNewLayout();
            chart.AddNewPlotVisOnly().val = 1;

            CT_PrintSettings printSettings = chartSpaceDocument.GetChartSpace().AddNewPrintSettings();
            printSettings.AddNewHeaderFooter();
            CT_PageMargins pageMargins = printSettings.AddNewPageMargins();
            pageMargins.b = 0.75;
            pageMargins.l = 0.70;
            pageMargins.r = 0.70;
            pageMargins.t = 0.75;
            pageMargins.header = 0.30;
            pageMargins.footer = 0.30;
            printSettings.AddNewPageSetup();
        }

        /**
         * Return the underlying CTChartSpace bean, the root element of the SpreadsheetML Chart part.
         *
         * @return the underlying CTChartSpace bean
         */

        internal CT_ChartSpace GetCTChartSpace()
        {
            return chartSpaceDocument.GetChartSpace();
        }

        /**
         * Return the underlying CTChart bean, within the Chart Space
         *
         * @return the underlying CTChart bean
         */

        internal CT_Chart GetCTChart()
        {
            return chart;
        }


        protected internal override void Commit()
        {
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);

            /*
               Saved chart space must have the following namespaces Set:
               <c:chartSpace
                  xmlns:c="http://schemas.Openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.Openxmlformats.org/drawingml/2006/main"
                  xmlns:r="http://schemas.Openxmlformats.org/officeDocument/2006/relationships">
             */
            PackagePart part = GetPackagePart();            
            chartSpaceDocument.Save(part.GetOutputStream());            
        }

        /**
         * Returns the parent graphic frame.
         * @return the graphic frame this chart belongs to
         */
        public XSSFGraphicFrame GetGraphicFrame()
        {
            return frame;
        }

        /**
         * Sets the parent graphic frame.
         */
        internal void SetGraphicFrame(XSSFGraphicFrame frame)
        {
            this.frame = frame;
        }

        public IChartDataFactory ChartDataFactory
        {
            get
            {
                return XSSFChartDataFactory.GetInstance();
            }
        }

        public IChartAxisFactory ChartAxisFactory
        {
            get
            {
                return this;
            }
        }

        public void Plot(IChartData data, params IChartAxis[] axis)
        {
            data.FillChart(this, axis);
        }

        public IValueAxis CreateValueAxis(AxisPosition pos)
        {
            long id = axis.Count + 1;
            XSSFValueAxis valueAxis = new XSSFValueAxis(this, id, pos);
            if (axis.Count == 1)
            {
                IChartAxis ax = axis[0];
                ax.CrossAxis(valueAxis);
                valueAxis.CrossAxis(ax);
            }
            axis.Add(valueAxis);
            return valueAxis;
        }
        public IChartAxis CreateCategoryAxis(AxisPosition pos)
        {
            long id = axis.Count + 1;
            XSSFCategoryAxis categoryAxis = new XSSFCategoryAxis(this, id, pos);
            if (axis.Count == 1)
            {
                IChartAxis ax = axis[0];
                ax.CrossAxis(categoryAxis);
                categoryAxis.CrossAxis(ax);
            }
            axis.Add(categoryAxis);
            return categoryAxis;
        }
        public List<IChartAxis> GetAxis()
        {
            if (axis.Count == 0 && HasAxis())
            {
                ParseAxis();
            }
            return axis;
        }

        public IManualLayout GetManualLayout()
        {
            return new XSSFManualLayout(this);
        }

        /**
         * @return true if only visible cells will be present on the chart,
         *         false otherwise
         */
        public bool IsPlotOnlyVisibleCells()
        {
            return chart.plotVisOnly.val==1?true:false;
        }

        /**
         * @param plotVisOnly a flag specifying if only visible cells should be
         *        present on the chart
         */
        public void SetPlotOnlyVisibleCells(bool plotVisOnly)
        {
            chart.plotVisOnly.val = plotVisOnly?1:0;
        }

        /**
         * Returns the title, or null if none is Set
         */
        public XSSFRichTextString Title
        {
            get
            {
                if (!chart.IsSetTitle())
                {
                    return null;
                }

                CT_Title title = chart.title;

                if (title.tx==null)
                    return null;
                if(title.tx.rich==null)
                    return null;
                return new XSSFRichTextString(title.tx.rich.ToString());
            }
        }

        public IChartLegend GetOrCreateLegend()
        {
            return new XSSFChartLegend(this);
        }

        public void DeleteLegend()
        {
            if (chart.IsSetLegend())
            {
                chart.unsetLegend();
            }
        }

        private bool HasAxis()
        {
            CT_PlotArea ctPlotArea = chart.plotArea;
            int totalAxisCount =
                (ctPlotArea.valAx == null ? 0 : ctPlotArea.valAx.Count) +
                (ctPlotArea.catAx == null ? 0 : ctPlotArea.catAx.Count) +
                (ctPlotArea.dateAx == null ? 0 : ctPlotArea.dateAx.Count) +
                (ctPlotArea.serAx == null ? 0 : ctPlotArea.serAx.Count);
            return totalAxisCount > 0;
        }

        private void ParseAxis()
        {
            ParseCategoryAxis();
            ParseValueAxis();
        }
        private void ParseCategoryAxis()
        {
            if (chart.plotArea.catAx == null)
                return;
            foreach (CT_CatAx catAx in chart.plotArea.catAx)
            {
                axis.Add(new XSSFCategoryAxis(this, catAx));
            }
        }
        private void ParseValueAxis()
        {
            if (chart.plotArea.valAx == null)
                return;
            foreach (CT_ValAx valAx in chart.plotArea.valAx)
            {
                axis.Add(new XSSFValueAxis(this, valAx));
            }
        }

    }
}



