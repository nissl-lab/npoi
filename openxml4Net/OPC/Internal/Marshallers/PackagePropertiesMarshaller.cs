using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;

namespace NPOI.OpenXml4Net.OPC.Internal.Marshallers
{
    /**
     * Package properties marshaller.
     *
     * @author CDubet, Julien Chable
     */
    public class PackagePropertiesMarshaller : PartMarshaller
    {

        private static string namespaceDC = PackagePropertiesPart.NAMESPACE_DC_URI;

        private static string namespaceCoreProperties = PackagePropertiesPart.NAMESPACE_CP_URI;

        private static string namespaceDcTerms = PackagePropertiesPart.NAMESPACE_DCTERMS_URI;

        private static string namespaceXSI = PackagePropertiesPart.NAMESPACE_XSI_URI;

        protected static String KEYWORD_CATEGORY = "category";

        protected static String KEYWORD_CONTENT_STATUS = "contentStatus";

        protected static String KEYWORD_CONTENT_TYPE = "contentType";

        protected static String KEYWORD_CREATED = "created";

        protected static String KEYWORD_CREATOR = "creator";

        protected static String KEYWORD_DESCRIPTION = "description";

        protected static String KEYWORD_IDENTIFIER = "identifier";

        protected static String KEYWORD_KEYWORDS = "keywords";

        protected static String KEYWORD_LANGUAGE = "language";

        protected static String KEYWORD_LAST_MODIFIED_BY = "lastModifiedBy";

        protected static String KEYWORD_LAST_PRINTED = "lastPrinted";

        protected static String KEYWORD_MODIFIED = "modified";

        protected static String KEYWORD_REVISION = "revision";

        protected static String KEYWORD_SUBJECT = "subject";

        protected static String KEYWORD_TITLE = "title";

        protected static String KEYWORD_VERSION = "version";

        PackagePropertiesPart propsPart;

        // The document
        protected XmlDocument xmlDoc = null;
        protected XmlNamespaceManager nsmgr = null;
        /**
         * Marshall package core properties to an XML document. Always return
         * <code>true</code>.
         */
        public virtual bool Marshall(PackagePart part, Stream out1)
        {
            if (!(part is PackagePropertiesPart))
                throw new ArgumentException(
                        "'part' must be a PackagePropertiesPart instance.");
            propsPart = (PackagePropertiesPart)part;

            // Configure the document
            xmlDoc = new XmlDocument();
            XmlElement rootElem = xmlDoc.CreateElement("coreProperties",namespaceCoreProperties);

            nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("cp",PackagePropertiesPart.NAMESPACE_CP_URI);
            nsmgr.AddNamespace("dc",PackagePropertiesPart.NAMESPACE_DC_URI);
            nsmgr.AddNamespace("dcterms",PackagePropertiesPart.NAMESPACE_DCTERMS_URI);
            nsmgr.AddNamespace("xsi", PackagePropertiesPart.NAMESPACE_XSI_URI);

            rootElem.SetAttribute("xmlns:cp", PackagePropertiesPart.NAMESPACE_CP_URI);
            rootElem.SetAttribute("xmlns:dc", PackagePropertiesPart.NAMESPACE_DC_URI);
            rootElem.SetAttribute("xmlns:dcterms", PackagePropertiesPart.NAMESPACE_DCTERMS_URI);
            rootElem.SetAttribute("xmlns:xsi", PackagePropertiesPart.NAMESPACE_XSI_URI);

            xmlDoc.AppendChild(rootElem);

            AddCategory();
            AddContentStatus();
            AddContentType();
            AddCreated();
            AddCreator();
            AddDescription();
            AddIdentifier();
            AddKeywords();
            AddLanguage();
            AddLastModifiedBy();
            AddLastPrinted();
            AddModified();
            AddRevision();
            AddSubject();
            AddTitle();
            AddVersion();
            return true;
        }

        /**
         * Add category property element if needed.
         */
        private void AddCategory()
        {
            if (propsPart.GetCategoryProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_CATEGORY, namespaceCoreProperties);
            XmlNode elem = null;
            if (elems.Count==0)
            {
                // Missing, we Add it
                elem =xmlDoc.CreateElement("cp", KEYWORD_CATEGORY, namespaceCoreProperties); 
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = elems[0];
                elem.InnerXml="";// clear the old value
            }
            elem.InnerText=propsPart.GetCategoryProperty();
        }

        /**
         * Add content status property element if needed.
         */
        private void AddContentStatus()
        {
            if (propsPart.GetContentStatusProperty()==null)
                return;
            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_CONTENT_STATUS, namespaceCoreProperties);
            XmlNode elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("cp", KEYWORD_CONTENT_STATUS, namespaceCoreProperties);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetContentStatusProperty();
        }

        /**
         * Add content type property element if needed.
         */
        private void AddContentType()
        {
            if (propsPart.GetContentTypeProperty()==null)
                return;
            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_CONTENT_TYPE,namespaceCoreProperties);
            XmlNode elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("cp", KEYWORD_CONTENT_TYPE, namespaceCoreProperties);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetContentTypeProperty();
        }

        /**
         * Add created property element if needed.
         */
        private void AddCreated()
        {
            if (propsPart.GetCreatedProperty() == null)
                return;
            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_CREATED,namespaceDcTerms);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dcterms", KEYWORD_CREATED, namespaceDcTerms);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.SetAttribute("type", namespaceXSI, "dcterms:W3CDTF");
            elem.InnerText = propsPart.GetCreatedPropertyString();
        }

        /**
         * Add creator property element if needed.
         */
        private void AddCreator()
        {
            if (propsPart.GetCreatorProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_CREATOR, namespaceDC);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dc", KEYWORD_CREATOR, namespaceDC);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }

            elem.InnerText = propsPart.GetCreatorProperty();
        }

        /**
         * Add description property element if needed.
         */
        private void AddDescription()
        {
            if (propsPart.GetDescriptionProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_DESCRIPTION, namespaceDC);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dc", KEYWORD_DESCRIPTION, namespaceDC);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }

            elem.InnerText = propsPart.GetDescriptionProperty();
        }

        /**
         * Add identifier property element if needed.
         */
        private void AddIdentifier()
        {
            if (propsPart.GetIdentifierProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_IDENTIFIER, namespaceDC);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dc", KEYWORD_IDENTIFIER, namespaceDC);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }

            elem.InnerText = propsPart.GetIdentifierProperty();
        }

        /**
         * Add keywords property element if needed.
         */
        private void AddKeywords()
        {
            if (propsPart.GetKeywordsProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_KEYWORDS, namespaceCoreProperties);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("cp", KEYWORD_KEYWORDS, namespaceCoreProperties);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetKeywordsProperty();
        }

        /**
         * Add language property element if needed.
         */
        private void AddLanguage()
        {
            if (propsPart.GetLanguageProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_LANGUAGE, namespaceDC);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dc", KEYWORD_LANGUAGE, namespaceDC);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetLanguageProperty();
        }

        /**
         * Add 'last modified by' property if needed.
         */
        private void AddLastModifiedBy()
        {
            if (propsPart.GetLastModifiedByProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_LAST_MODIFIED_BY, namespaceCoreProperties);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("cp", KEYWORD_LAST_MODIFIED_BY, namespaceCoreProperties);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetLastModifiedByProperty();
        }

        /**
         * Add 'last printed' property if needed.
         *
         */
        private void AddLastPrinted()
        {
            if (propsPart.GetLastPrintedProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_LAST_PRINTED, namespaceCoreProperties);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("cp", KEYWORD_LAST_PRINTED, namespaceCoreProperties);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetLastPrintedPropertyString();
        }

        /**
         * Add modified property element if needed.
         */
        private void AddModified()
        {
            if (propsPart.GetModifiedProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_MODIFIED, namespaceDcTerms);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dcterms", KEYWORD_MODIFIED, namespaceDcTerms);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetModifiedPropertyString();
            elem.SetAttribute("type", namespaceXSI, "dcterms:W3CDTF");
        }

        /**
         * Add revision property if needed.
         */
        private void AddRevision()
        {
            if (propsPart.GetRevisionProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_REVISION, namespaceCoreProperties);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("cp", KEYWORD_REVISION, namespaceCoreProperties);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetRevisionProperty();
        }

        /**
         * Add subject property if needed.
         */
        private void AddSubject()
        {
            if (propsPart.GetSubjectProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_SUBJECT, namespaceDC);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dc", KEYWORD_SUBJECT, namespaceDC);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetSubjectProperty();
        }

        /**
         * Add title property if needed.
         */
        private void AddTitle()
        {
            if (propsPart.GetTitleProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_TITLE, namespaceDC);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("dc", KEYWORD_TITLE, namespaceDC);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetTitleProperty();
        }

        private void AddVersion()
        {
            if (propsPart.GetVersionProperty() == null)
                return;

            XmlNodeList elems = xmlDoc.DocumentElement.GetElementsByTagName(KEYWORD_VERSION, namespaceCoreProperties);
            XmlElement elem = null;
            if (elems.Count == 0)
            {
                // Missing, we Add it
                elem = xmlDoc.CreateElement("cp", KEYWORD_VERSION, namespaceCoreProperties);
                xmlDoc.DocumentElement.AppendChild(elem);
            }
            else
            {
                elem = (XmlElement)elems[0];
                elem.InnerXml = "";// clear the old value
            }
            elem.InnerText = propsPart.GetVersionProperty();
        }
    }
}