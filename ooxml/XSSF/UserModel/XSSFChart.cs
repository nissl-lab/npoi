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

namespace NPOI.xssf.usermodel;








using javax.xml.namespace.QName;

using NPOI.POIXMLDocumentPart;
using NPOI.Openxml4j.opc.PackagePart;
using NPOI.Openxml4j.opc.PackageRelationship;
using NPOI.util.Internal;
using NPOI.ss.usermodel.Chart;
using NPOI.ss.usermodel.charts.ChartAxis;
using NPOI.ss.usermodel.charts.ChartAxisFactory;
using NPOI.xssf.usermodel.charts.XSSFChartDataFactory;
using NPOI.xssf.usermodel.charts.XSSFChartAxis;
using NPOI.xssf.usermodel.charts.XSSFValueAxis;
using NPOI.xssf.usermodel.charts.XSSFManualLayout;
using NPOI.xssf.usermodel.charts.XSSFChartLegend;
using NPOI.ss.usermodel.charts.ChartData;
using NPOI.ss.usermodel.charts.AxisPosition;
using org.apache.xmlbeans.XmlException;
using org.apache.xmlbeans.XmlObject;
using org.apache.xmlbeans.XmlOptions;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTChart;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTChartSpace;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTTitle;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.ChartSpaceDocument;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTLayout;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTManualLayout;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTPlotArea;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTValAx;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTPrintSettings;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.CTPageMargins;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.STLayoutTarget;
using org.Openxmlformats.schemas.Drawingml.x2006.chart.STLayoutMode;
using org.Openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;
using org.w3c.dom.NodeList;
using org.w3c.dom.Text;

/**
 * Represents a SpreadsheetML Chart
 * @author Nick Burch
 * @author Roman Kashitsyn
 */
public class XSSFChart : POIXMLDocumentPart : Chart, ChartAxisFactory {

	/**
	 * Parent graphic frame.
	 */
	private XSSFGraphicFrame frame;

	/**
	 * Root element of the SpreadsheetML Chart part
	 */
	private CTChartSpace chartSpace;
	/**
	 * The Chart within that
	 */
	private CTChart chart;

	List<XSSFChartAxis> axis;

	/**
	 * Create a new SpreadsheetML chart
	 */
	protected XSSFChart() {
		base();
		axis = new List<XSSFChartAxis>();
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
	protected XSSFChart(PackagePart part, PackageRelationship rel) , XmlException {
		base(part, rel);

		chartSpace = ChartSpaceDocument.Factory.Parse(part.GetInputStream()).GetChartSpace(); 
		chart = chartSpace.GetChart();
	}

	/**
	 * Construct a new CTChartSpace bean.
	 * By default, it's just an empty placeholder for chart objects.
	 *
	 * @return a new CTChartSpace bean
	 */
	private void CreateChart() {
		chartSpace = CTChartSpace.Factory.newInstance();
		chart = chartSpace.AddNewChart();
		CTPlotArea plotArea = chart.AddNewPlotArea();

		plotArea.AddNewLayout();
		chart.AddNewPlotVisOnly().SetVal(true);

		CTPrintSettings printSettings = chartSpace.AddNewPrintSettings();
		printSettings.AddNewHeaderFooter();

		CTPageMargins pageMargins = printSettings.AddNewPageMargins();
		pageMargins.SetB(0.75);
		pageMargins.SetL(0.70);
		pageMargins.SetR(0.70);
		pageMargins.SetT(0.75);
		pageMargins.SetHeader(0.30);
		pageMargins.SetFooter(0.30);
		printSettings.AddNewPageSetup();
	}

	/**
	 * Return the underlying CTChartSpace bean, the root element of the SpreadsheetML Chart part.
	 *
	 * @return the underlying CTChartSpace bean
	 */
	
	public CTChartSpace GetCTChartSpace(){
		return chartSpace;
	}

	/**
	 * Return the underlying CTChart bean, within the Chart Space
	 *
	 * @return the underlying CTChart bean
	 */
	
	public CTChart GetCTChart(){
		return chart;
	}

	
	protected void Commit()  {
		XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);

		/*
		   Saved chart space must have the following namespaces Set:
		   <c:chartSpace
		      xmlns:c="http://schemas.Openxmlformats.org/drawingml/2006/chart"
		      xmlns:a="http://schemas.Openxmlformats.org/drawingml/2006/main"
		      xmlns:r="http://schemas.Openxmlformats.org/officeDocument/2006/relationships">
		 */
		xmlOptions.SetSaveSyntheticDocumentElement(new QName(CTChartSpace.type.GetName().GetNamespaceURI(), "chartSpace", "c"));
		Dictionary<String, String> map = new Dictionary<String, String>();
		map.Put(XSSFDrawing.NAMESPACE_A, "a");
		map.Put(XSSFDrawing.NAMESPACE_C, "c");
		map.Put(STRelationshipId.type.GetName().GetNamespaceURI(), "r");
		xmlOptions.SetSaveSuggestedPrefixes(map);

		PackagePart part = GetPackagePart();
		OutputStream out = part.GetOutputStream();
		chartSpace.save(out, xmlOptions);
		out.Close();
	}

	/**
	 * Returns the parent graphic frame.
	 * @return the graphic frame this chart belongs to
	 */
	public XSSFGraphicFrame GetGraphicFrame() {
		return frame;
	}

	/**
	 * Sets the parent graphic frame.
	 */
	protected void SetGraphicFrame(XSSFGraphicFrame frame) {
		this.frame = frame;
	}

	public XSSFChartDataFactory GetChartDataFactory() {
		return XSSFChartDataFactory.GetInstance();
	}

	public XSSFChart GetChartAxisFactory() {
		return this;
	}

	public void plot(ChartData data, ChartAxis... axis) {
		data.FillChart(this, axis);
	}

	public XSSFValueAxis CreateValueAxis(AxisPosition pos) {
		long id = axis.Count + 1;
		XSSFValueAxis valueAxis = new XSSFValueAxis(this, id, pos);
		if (axis.Count == 1) {
			ChartAxis ax = axis.Get(0);
			ax.crossAxis(valueAxis);
			valueAxis.crossAxis(ax);
		}
		axis.Add(valueAxis);
		return valueAxis;
	}

	public List<? : XSSFChartAxis> GetAxis() {
		if (axis.IsEmpty() && hasAxis()) {
			ParseAxis();
		}
		return axis;
	}

	public XSSFManualLayout GetManualLayout() {
		return new XSSFManualLayout(this);
	}

	/**
	 * @return true if only visible cells will be present on the chart,
	 *         false otherwise
	 */
	public bool IsPlotOnlyVisibleCells() {
		return chart.GetPlotVisOnly().GetVal();
	}

	/**
	 * @param plotVisOnly a flag specifying if only visible cells should be
	 *        present on the chart
	 */
	public void SetPlotOnlyVisibleCells(bool plotVisOnly) {
		chart.GetPlotVisOnly().SetVal(plotVisOnly);
	}

	/**
	 * Returns the title, or null if none is Set
	 */
	public XSSFRichTextString GetTitle() {
		if(! chart.IsSetTitle()) {
			return null;
		}

		// TODO Do properly
		CTTitle title = chart.GetTitle();

		StringBuilder text = new StringBuilder();
		XmlObject[] t = title
			.selectPath("declare namespace a='"+XSSFDrawing.NAMESPACE_A+"' .//a:t");
		for (int m = 0; m < t.Length; m++) {
			NodeList kids = t[m].GetDomNode().GetChildNodes();
			for (int n = 0; n < kids.GetLength(); n++) {
				if (kids.item(n) is Text) {
					text.Append(kids.item(n).GetNodeValue());
				}
			}
		}

		return new XSSFRichTextString(text.ToString());
	}

	public XSSFChartLegend GetOrCreateLegend() {
		return new XSSFChartLegend(this);
	}

	public void deleteLegend() {
		if (chart.IsSetLegend()) {
			chart.unsetLegend();
		}
	}

	private bool HasAxis() {
		CTPlotArea ctPlotArea = chart.GetPlotArea();
		int totalAxisCount =
			ctPlotArea.sizeOfValAxArray()  +
			ctPlotArea.sizeOfCatAxArray()  +
			ctPlotArea.sizeOfDateAxArray() +
			ctPlotArea.sizeOfSerAxArray();
		return totalAxisCount > 0;
	}

	private void ParseAxis() {
		// TODO: add other axis types
		ParseValueAxis();
	}

	private void ParseValueAxis() {
		foreach (CTValAx valAx in chart.GetPlotArea().GetValAxArray()) {
			axis.Add(new XSSFValueAxis(this, valAx));
		}
	}

}



