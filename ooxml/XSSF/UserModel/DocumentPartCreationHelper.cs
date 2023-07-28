/*
 *  ====================================================================
 *    Licensed to the collaborators of the NPOI project under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The collaborators licenses this file to You under the Apache License, Version 2.0
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
using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.Model;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.XSSF.UserModel
{
    internal static class DocumentPartCreationHelper
    {
        public static POIXMLDocumentPart CreateDocumentPart(Type cls, Type[] types, Object[] values)
        {
            if ((values is null || values.Length == 0) && (types is null || types.Length == 0))
            {
                return CreateDocumentPartParamless(cls);
            }

            if (values.Length == 1 && types.Length == 1 && types[0] == typeof(PackagePart))
            {
                return CreateDocumentPartOneParam(cls, (PackagePart)values[1]);
            }

            if (values.Length == 2 && types.Length == 2 && types[0] == typeof(POIXMLDocumentPart) && types[1] == typeof(PackagePart))
            {
                return CreateDocumentPartTwoParams(cls, (POIXMLDocumentPart)values[0], (PackagePart)values[1]);
            }

            ThrowHelper_MissingMethod();
            return null;
        }

        private static void ThrowHelper_MissingMethod()
        {
            throw new MissingMethodException();
        }
        private static POIXMLDocumentPart CreateDocumentPartParamless(Type cls)
        {
            if (cls == typeof(XWPFDocument))
                return new XWPFDocument();
            if (cls == typeof(XSSFWorkbook))
                return new XSSFWorkbook();
            if (cls == typeof(XWPFComments))
                return new XWPFComments();
            if (cls == typeof(XWPFFootnotes))
                return new XWPFFootnotes();
            if (cls == typeof(XWPFFooter))
                return new XWPFFooter();
            if (cls == typeof(XWPFHeader))
                return new XWPFHeader();
            if (cls == typeof(XWPFNumbering))
                return new XWPFNumbering();
            if (cls == typeof(XWPFPictureData))
                return XWPFPictureData.InternalCreateInstance();
            if (cls == typeof(XWPFSettings))
                return new XWPFSettings();
            if (cls == typeof(XWPFStyles))
                return new XWPFStyles();
            if (cls == typeof(XSSFChart))
                return new XSSFChart();
            if (cls == typeof(XSSFDrawing))
                return new XSSFDrawing();
            if (cls == typeof(XSSFPictureData))
                return new XSSFPictureData();
            if (cls == typeof(XSSFPivotCache))
                return new XSSFPivotCache();
            if (cls == typeof(XSSFPivotCacheDefinition))
                return new XSSFPivotCacheDefinition();
            if (cls == typeof(XSSFPivotCacheRecords))
                return new XSSFPivotCacheRecords();
            if (cls == typeof(XSSFPivotTable))
                return new XSSFPivotTable();
            if (cls == typeof(XSSFSheet))
                return new XSSFSheet();
            //if (cls == typeof(XSSFChartSheet))
                //return new XSSFChartSheet();
            //if (cls == typeof(XSSFDialogsheet))
                //return new XSSFDialogsheet();
            if (cls == typeof(XSSFTable))
                return new XSSFTable();
            if (cls == typeof(XSSFVBAPart))
                return new XSSFVBAPart();
            if (cls == typeof(XSSFVMLDrawing))
                return new XSSFVMLDrawing();
            if (cls == typeof(CalculationChain))
                return new CalculationChain();
            if (cls == typeof(CommentsTable))
                return new CommentsTable();
            if (cls == typeof(ExternalLinksTable))
                return new ExternalLinksTable();
            if (cls == typeof(MapInfo))
                return new MapInfo();
            if (cls == typeof(SharedStringsTable))
                return new SharedStringsTable();
            if (cls == typeof(SingleXmlCells))
                return new SingleXmlCells();
            if (cls == typeof(StylesTable))
                return new StylesTable();
            if (cls == typeof(ThemesTable))
                return new ThemesTable();

            ThrowHelper_MissingMethod();
            return null;
        }
        private static POIXMLDocumentPart CreateDocumentPartOneParam(Type cls, PackagePart part)
        {
            //if (cls == typeof(XWPFDocument))
                //return new XWPFDocument(part);
            //if (cls == typeof(XSSFWorkbook))
                //return new XSSFWorkbook(part);
            //if (cls == typeof(XWPFComments))
                //return new XWPFComments(part);
            if (cls == typeof(XWPFFootnotes))
                return new XWPFFootnotes(part);
            //if (cls == typeof(XWPFFooter))
                //return new XWPFFooter(part);
            //if (cls == typeof(XWPFHeader))
                //return new XWPFHeader(part);
            if (cls == typeof(XWPFNumbering))
                return new XWPFNumbering(part);
            if (cls == typeof(XWPFPictureData))
                return new XWPFPictureData(part);
            if (cls == typeof(XWPFSettings))
                return new XWPFSettings(part);
            if (cls == typeof(XWPFStyles))
                return new XWPFStyles(part);
            if (cls == typeof(XSSFChart))
                return XSSFChart.InternalCreateInstance(part);
            if (cls == typeof(XSSFDrawing))
                return new XSSFDrawing(part);
            if (cls == typeof(XSSFPictureData))
                return new XSSFPictureData(part);
            if (cls == typeof(XSSFPivotCache))
                return XSSFPivotCache.InternalCreateInstance(part);
            if (cls == typeof(XSSFPivotCacheDefinition))
                return XSSFPivotCacheDefinition.InternalCreateInstance(part);
            if (cls == typeof(XSSFPivotCacheRecords))
                return XSSFPivotCacheRecords.InternalCreateInstance(part);
            if (cls == typeof(XSSFPivotTable))
                return XSSFPivotTable.InternalCreateInstance(part);
            if (cls == typeof(XSSFSheet))
                return new XSSFSheet(part);
            if (cls == typeof(XSSFChartSheet))
                return XSSFChartSheet.InternalCreateInstance(part);
            //if (cls == typeof(XSSFDialogsheet))
                //return new XSSFDialogsheet(part);
            if (cls == typeof(XSSFTable))
                return new XSSFTable(part);
            if (cls == typeof(XSSFVBAPart))
                return new XSSFVBAPart(part);
            if (cls == typeof(XSSFVMLDrawing))
                return XSSFVMLDrawing.InternalCreateInstance(part);
            if (cls == typeof(CalculationChain))
                return new CalculationChain(part);
            if (cls == typeof(CommentsTable))
                return new CommentsTable(part);
            if (cls == typeof(ExternalLinksTable))
                return new ExternalLinksTable(part);
            if (cls == typeof(MapInfo))
                return new MapInfo(part);
            if (cls == typeof(SharedStringsTable))
                return new SharedStringsTable(part);
            if (cls == typeof(SingleXmlCells))
                return new SingleXmlCells(part);
            if (cls == typeof(StylesTable))
                return new StylesTable(part);
            if (cls == typeof(ThemesTable))
                return new ThemesTable(part);

            ThrowHelper_MissingMethod();
            return null;
        }
        private static POIXMLDocumentPart CreateDocumentPartTwoParams(Type cls, POIXMLDocumentPart parent, PackagePart part)
        {
            //if (cls == typeof(XWPFDocument))
                //return new XWPFDocument(parent, part);
            //if (cls == typeof(XSSFWorkbook))
                //return new XSSFWorkbook(parent, part);
            if (cls == typeof(XWPFComments))
                return new XWPFComments(parent, part);
            //if (cls == typeof(XWPFFootnotes))
                //return new XWPFFootnotes(parent, part);
            if (cls == typeof(XWPFFooter))
                return new XWPFFooter(parent, part);
            if (cls == typeof(XWPFHeader))
                return new XWPFHeader(parent, part);
            //if (cls == typeof(XWPFNumbering))
                //return new XWPFNumbering(parent, part);
            //if (cls == typeof(XWPFPictureData))
                //return new XWPFPictureData(parent, part);
            //if (cls == typeof(XWPFSettings))
                //return new XWPFSettings(parent, part);
            //if (cls == typeof(XWPFStyles))
                //return new XWPFStyles(parent, part);
            //if (cls == typeof(XSSFChart))
                //return new XSSFChart(parent, part);
            //if (cls == typeof(XSSFDrawing))
                //return new XSSFDrawing(parent, part);
            //if (cls == typeof(XSSFPictureData))
                //return new XSSFPictureData(parent, part);
            //if (cls == typeof(XSSFPivotCache))
                //return new XSSFPivotCache(parent, part);
            //if (cls == typeof(XSSFPivotCacheDefinition))
                //return new XSSFPivotCacheDefinition(parent, part);
            //if (cls == typeof(XSSFPivotCacheRecords))
                //return new XSSFPivotCacheRecords(parent, part);
            //if (cls == typeof(XSSFPivotTable))
                //return new XSSFPivotTable(parent, part);
            //if (cls == typeof(XSSFSheet))
                //return new XSSFSheet(parent, part);
            //if (cls == typeof(XSSFChartSheet))
                //return new XSSFChartSheet(parent, part);
            //if (cls == typeof(XSSFDialogsheet))
                //return new XSSFDialogsheet(parent, part);
            //if (cls == typeof(XSSFTable))
                //return new XSSFTable(parent, part);
            //if (cls == typeof(XSSFVBAPart))
                //return new XSSFVBAPart(parent, part);
            //if (cls == typeof(XSSFVMLDrawing))
                //return new XSSFVMLDrawing(parent, part);
            //if (cls == typeof(CalculationChain))
                //return new CalculationChain(parent, part);
            //if (cls == typeof(CommentsTable))
                //return new CommentsTable(parent, part);
            //if (cls == typeof(ExternalLinksTable))
                //return new ExternalLinksTable(parent, part);
            //if (cls == typeof(MapInfo))
                //return new MapInfo(parent, part);
            //if (cls == typeof(SharedStringsTable))
                //return new SharedStringsTable(parent, part);
            //if (cls == typeof(SingleXmlCells))
                //return new SingleXmlCells(parent, part);
            //if (cls == typeof(StylesTable))
                //return new StylesTable(parent, part);
            //if (cls == typeof(ThemesTable))
                //return new ThemesTable(parent, part);

            ThrowHelper_MissingMethod();
            return null;
        }
    }
}
