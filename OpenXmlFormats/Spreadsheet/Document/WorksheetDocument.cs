﻿using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class WorksheetDocument
    {
        CT_Worksheet sheet = null;

        public WorksheetDocument()
        {
        }
        public WorksheetDocument(CT_Worksheet sheet)
        {
            this.sheet = sheet;
        }
        public static WorksheetDocument Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceMgr)
        {
            CT_Worksheet obj = CT_Worksheet.Parse(xmldoc.DocumentElement, namespaceMgr);
            return new WorksheetDocument(obj);
        }
        public CT_Worksheet GetWorksheet()
        {
            return sheet;
        }
        public void SetChartsheet(CT_Worksheet sheet)
        {
            this.sheet = sheet;
        }
        public void Save(Stream stream, bool leaveOpen)
        {
            this.sheet.Write(stream, leaveOpen);
        }

        public async Task SaveAsync(Stream stream, bool leaveOpen, CancellationToken cancellationToken = default)
        {
            await this.sheet.WriteAsync(stream, leaveOpen, cancellationToken).ConfigureAwait(false);
        }
    }
}
