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

        /// <summary>
        /// Returns the title, or null if none is Set
        /// </summary>
        public XSSFRichTextString TitleText
        {
            get
            {
                if(!chart.IsSetTitle())
                {
                    return null;
                }

                CT_Title title = chart.title;

                if(title.tx==null)
                    return null;
                if(title.tx.rich!=null)
                    return new XSSFRichTextString(title.tx.rich.ToString());
                else
                    return new XSSFRichTextString("");
            }
        }

        /// <summary>
        /// Get the chart title formula expression if there is one
        /// </summary>
        public String TitleFormula
        {
            get
            {
                if(!GetCTChart().IsSetTitle())
                {
                    return null;
                }

                var title = GetCTChart().title;

                if(!title.IsSetTx())
                {
                    return null;
                }

                var tx = title.tx;

                if(!tx.IsSetStrRef())
                {
                    return null;
                }

                return tx.strRef.f;
            }
        }
        /// <summary>
        /// Set the formula expression to use for the chart title
        /// </summary>
        /// <param name="formula"></param>
        public void SetTitleFormula(String formula)
        {
            CT_Title ctTitle;
            if(GetCTChart().IsSetTitle())
            {
                ctTitle = GetCTChart().title;
            }
            else
            {
                ctTitle = GetCTChart().AddNewTitle();
            }

            CT_Tx tx;
            if(ctTitle.IsSetTx())
            {
                tx = ctTitle.tx;
            }
            else
            {
                tx = ctTitle.AddNewTx();
            }

            if(tx.IsSetRich())
            {
                tx.UnsetRich();
            }

            CT_StrRef strRef;
            if(tx.IsSetStrRef())
            {
                strRef = tx.strRef;
            }
            else
            {
                strRef = tx.AddNewStrRef();
            }

            strRef.f = formula;
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



