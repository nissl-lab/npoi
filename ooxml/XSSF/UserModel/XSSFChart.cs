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
    using NPOI.OpenXmlFormats.Dml.Chart;
    using System.Collections.Generic;
    using NPOI.OpenXml4Net.OPC;
    using System;
    using System.IO;
    using System.Text;
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.XDDF.UserModel.Chart;

    /**
     * Represents a SpreadsheetML Chart
     * @author Nick Burch
     * @author Roman Kashitsyn
     */
    public class XSSFChart : XDDFChart
    {

        /**
         * Parent graphic frame.
         */
        private XSSFGraphicFrame frame;

        /**
         * Create a new SpreadsheetML chart
         */
        public XSSFChart()
            : base()
        {
            CreateChart();
        }

        /**
         * Construct a SpreadsheetML chart from a package part.
         *
         * @param part the package part holding the chart data,
         * the content type must be <code>application/vnd.openxmlformats-officedocument.drawingml.chart+xml</code>
         * @param rel  the package relationship holding this chart,
         * the relationship type must be http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart
         * 
         * @since POI 3.14-Beta1
         */
        protected XSSFChart(PackagePart part)
            : base(part)
        {
        }
        /**
         * Construct a new CTChartSpace bean.
         * By default, it's just an empty placeholder for chart objects.
         *
         * @return a new CTChartSpace bean
         */
        private void CreateChart()
        {
            CT_PlotArea plotArea = chart.AddNewPlotArea();

            plotArea.AddNewLayout();
            chart.AddNewPlotVisOnly().val = 1;

            CT_PrintSettings printSettings = chartSpace.AddNewPrintSettings();
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
            chartSpace.Save(part.GetOutputStream());
        }
        /// <summary>
        /// Returns the parent graphic frame.
        /// </summary>
        /// <returns></returns>
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

        public void SetCTDispBlanksAs(CT_DispBlanksAs disp)
        {
            chart.dispBlanksAs = disp;
        }

        protected override POIXMLRelation GetChartRelation()
        {
            return XSSFRelation.CHART;
        }

        protected override POIXMLRelation GetChartWorkbookRelation()
        {
            return XSSFRelation.CHART;
        }

        protected override POIXMLFactory GetChartFactory()
        {
            return XSSFFactory.GetInstance();
        }
    }
}



