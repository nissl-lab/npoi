using System.IO;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Document wrapper around the <c>ThreadedComments</c> part of a worksheet.
    /// </summary>
    public class ThreadedCommentsDocument
    {
        private CT_ThreadedComments threadedComments = null;

        public ThreadedCommentsDocument()
        {
        }

        public ThreadedCommentsDocument(CT_ThreadedComments threadedComments)
        {
            this.threadedComments = threadedComments;
        }

        public static ThreadedCommentsDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager namespaceManager)
        {
            ThreadedCommentsDocument doc = new ThreadedCommentsDocument();
            doc.threadedComments = CT_ThreadedComments.Parse(xmlDoc.DocumentElement, namespaceManager);
            return doc;
        }

        public CT_ThreadedComments GetThreadedComments()
        {
            return threadedComments;
        }

        public void SetThreadedComments(CT_ThreadedComments threadedComments)
        {
            this.threadedComments = threadedComments;
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                this.threadedComments.Write(sw);
            }
        }
    }
}
