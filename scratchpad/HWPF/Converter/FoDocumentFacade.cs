/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
namespace NPOI.HWPF.Converter
{
    using System;
    using System.Collections.Generic;
    using System.Xml;

    public class FoDocumentFacade
    {
        private const String NS_RDF = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";

        private const String NS_XSLFO = "http://www.w3.org/1999/XSL/Format";

        protected XmlElement declarations;
        protected XmlDocument document;
        protected XmlElement layoutMasterSet;
        protected XmlElement propertiesRoot;
        protected XmlElement root;

        public FoDocumentFacade(XmlDocument document)
        {
            this.document = document;
            root = document.CreateElement("fo:root", NS_XSLFO);
            document.AppendChild(root);

            layoutMasterSet = document.CreateElement("fo:layout-master-set", NS_XSLFO);
            root.AppendChild(layoutMasterSet);

            declarations = document.CreateElement("fo:declarations", NS_XSLFO);
            root.AppendChild(declarations);
        }

        public XmlElement AddFlowToPageSequence(XmlElement pageSequence,
                String flowName)
        {
            XmlElement flow = document.CreateElement("fo:flow", NS_XSLFO);
            flow.SetAttribute("flow-name", flowName);
            pageSequence.AppendChild(flow);

            return flow;
        }

        public XmlElement AddListItem(XmlElement listBlock)
        {
            XmlElement result = CreateListItem();
            listBlock.AppendChild(result);
            return result;
        }

        public XmlElement AddListItemBody(XmlElement listItem)
        {
            XmlElement result = CreateListItemBody();
            listItem.AppendChild(result);
            return result;
        }

        public XmlElement AddListItemLabel(XmlElement listItem, String text)
        {
            XmlElement result = CreateListItemLabel(text);
            listItem.AppendChild(result);
            return result;
        }

        public XmlElement AddPageSequence(String pageMaster)
        {
            XmlElement pageSequence = document.CreateElement("fo:page-sequence", NS_XSLFO);
            pageSequence.SetAttribute("master-reference", pageMaster);
            root.AppendChild(pageSequence);
            return pageSequence;
        }

        public XmlElement AddRegionBody(XmlElement pageMaster)
        {
            XmlElement regionBody = document.CreateElement(
                   "fo:region-body", NS_XSLFO);
            pageMaster.AppendChild(regionBody);

            return regionBody;
        }

        public XmlElement AddSimplePageMaster(String masterName)
        {
            XmlElement simplePageMaster = document.CreateElement(
                   "fo:simple-page-master", NS_XSLFO);
            simplePageMaster.SetAttribute("master-name", masterName);
            layoutMasterSet.AppendChild(simplePageMaster);

            return simplePageMaster;
        }

        public XmlElement CreateBasicLinkExternal(String externalDestination)
        {
            XmlElement basicLink = document.CreateElement(
                   "fo:basic-link", NS_XSLFO);
            basicLink.SetAttribute("external-destination", externalDestination);
            return basicLink;
        }

        public XmlElement CreateBasicLinkInternal(String internalDestination)
        {
            XmlElement basicLink = document.CreateElement(
                   "fo:basic-link", NS_XSLFO);
            basicLink.SetAttribute("internal-destination", internalDestination);
            return basicLink;
        }

        public XmlElement CreateBlock()
        {
            return document.CreateElement("fo:block", NS_XSLFO);
        }

        public XmlElement CreateExternalGraphic(String source)
        {
            XmlElement result = document.CreateElement(
                    "fo:external-graphic", NS_XSLFO);
            result.SetAttribute("src", "url('" + source + "')");
            return result;
        }

        public XmlElement CreateFootnote()
        {
            return document.CreateElement("fo:footnote", NS_XSLFO);
        }

        public XmlElement CreateFootnoteBody()
        {
            return document.CreateElement("fo:footnote-body", NS_XSLFO);
        }

        public XmlElement CreateInline()
        {
            return document.CreateElement("fo:inline", NS_XSLFO);
        }

        public XmlElement CreateLeader()
        {
            return document.CreateElement("fo:leader", NS_XSLFO);
        }

        public XmlElement CreateListBlock()
        {
            return document.CreateElement("fo:list-block", NS_XSLFO);
        }

        public XmlElement CreateListItem()
        {
            return document.CreateElement("fo:list-item", NS_XSLFO);
        }

        public XmlElement CreateListItemBody()
        {
            return document.CreateElement("fo:list-item-body", NS_XSLFO);
        }

        public XmlElement CreateListItemLabel(String text)
        {
            XmlElement result = document.CreateElement(
                    "fo:list-item-label", NS_XSLFO);
            XmlElement block = CreateBlock();
            block.AppendChild(document.CreateTextNode(text));
            result.AppendChild(block);
            return result;
        }

        public XmlElement CreateTable()
        {
            return document.CreateElement("fo:table", NS_XSLFO);
        }

        public XmlElement CreateTableBody()
        {
            return document.CreateElement("fo:table-body", NS_XSLFO);
        }

        public XmlElement CreateTableCell()
        {
            return document.CreateElement("fo:table-cell", NS_XSLFO);
        }

        public XmlElement CreateTableHeader()
        {
            return document.CreateElement("fo:table-header", NS_XSLFO);
        }

        public XmlElement CreateTableRow()
        {
            return document.CreateElement("fo:table-row", NS_XSLFO);
        }

        public XmlText CreateText(String data)
        {
            return document.CreateTextNode(data);
        }

        public XmlDocument Document
        {
            get
            {
                return document;
            }
        }

        protected XmlElement GetOrCreatePropertiesRoot()
        {
            if (propertiesRoot != null)
                return propertiesRoot;

            // See http://xmlgraphics.apache.org/fop/0.95/metadata.html

            XmlElement xmpmeta = document.CreateElement("adobe:ns:meta/",
                    "x:xmpmeta", NS_XSLFO);
            declarations.AppendChild(xmpmeta);

            XmlElement rdf = document.CreateElement("rdf:RDF", NS_RDF);
            xmpmeta.AppendChild(rdf);

            propertiesRoot = document.CreateElement("rdf:Description", NS_RDF);
            propertiesRoot.SetAttribute("rdf:about", NS_RDF, "");
            rdf.AppendChild(propertiesRoot);

            return propertiesRoot;
        }

        public void SetCreator(String value)
        {
            SetDublinCoreProperty("creator", value);
        }

        public void SetCreatorTool(String value)
        {
            SetXmpProperty("CreatorTool", value);
        }

        public void SetDescription(String value)
        {
            XmlElement element = SetDublinCoreProperty("description", value);

            if (element != null)
            {
                element.SetAttribute("xml:lang", "http://www.w3.org/XML/1998/namespace", "x-default");
            }
        }

        public XmlElement SetDublinCoreProperty(String name, String value)
        {
            return SetProperty("http://purl.org/dc/elements/1.1/", "dc", name, value);
        }

        public void SetKeywords(String value)
        {
            SetPdfProperty("Keywords", value);
        }

        public XmlElement SetPdfProperty(String name, String value)
        {
            return SetProperty("http://ns.adobe.com/pdf/1.3/", "pdf", name, value);
        }

        public void SetProducer(String value)
        {
            SetPdfProperty("Producer", value);
        }

        protected XmlElement SetProperty(String ns, String prefix,
                String name, String value)
        {
            XmlElement propertiesRoot = GetOrCreatePropertiesRoot();
            XmlNodeList existingChildren = propertiesRoot.ChildNodes;
            for (int i = 0; i < existingChildren.Count; i++)
            {
                XmlNode child = existingChildren[i];
                if (child.NodeType == XmlNodeType.Element)
                {
                    XmlElement childElement = (XmlElement)child;
                    if (!string.IsNullOrEmpty(childElement.NamespaceURI)
                            && !string.IsNullOrEmpty(childElement.LocalName)
                            && ns.Equals(childElement.NamespaceURI)
                            && name.Equals(childElement.LocalName))
                    {
                        propertiesRoot.RemoveChild(childElement);
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(value))
            {
                XmlElement property = document.CreateElement(ns, prefix + ":" + name);
                property.AppendChild(document.CreateTextNode(value));
                propertiesRoot.AppendChild(property);
                return property;
            }

            return null;
        }

        public void SetSubject(String value)
        {
            SetDublinCoreProperty("title", value);
        }

        public void SetTitle(String value)
        {
            SetDublinCoreProperty("title", value);
        }

        public XmlElement SetXmpProperty(String name, String value)
        {
            return SetProperty("http://ns.adobe.com/xap/1.0/", "xmp", name, value);
        }

    }
}