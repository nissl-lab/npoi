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
    using System.Xml;
    using NPOI.Util;

    public class TextDocumentFacade
    {
        protected XmlElement body;
        protected XmlDocument document;
        protected XmlElement head;
        protected XmlElement root;

        protected XmlElement title;
        protected XmlText titleText;

        public TextDocumentFacade(XmlDocument document)
        {
            this.document = document;

            root = document.CreateElement("html");
            document.AppendChild(root);

            body = document.CreateElement("body");
            head = document.CreateElement("head");

            root.AppendChild(head);
            root.AppendChild(body);

            title = document.CreateElement("title");
            titleText = document.CreateTextNode("");
            head.AppendChild(title);
        }

        public void AddAuthor(String value)
        {
            AddMeta("Author", value);
        }

        public void AddDescription(String value)
        {
            AddMeta("Description", value);
        }

        public void AddKeywords(String value)
        {
            AddMeta("Keywords", value);
        }

        public void AddMeta(String name, String value)
        {
            XmlElement meta = document.CreateElement("meta");

            XmlElement metaName = document.CreateElement("name");
            metaName.AppendChild(document.CreateTextNode(name + ": "));
            meta.AppendChild(metaName);

            XmlElement metaValue = document.CreateElement("value");
            metaValue.AppendChild(document.CreateTextNode(value + "\n"));
            meta.AppendChild(metaValue);

            head.AppendChild(meta);
        }

        public XmlElement CreateBlock()
        {
            return document.CreateElement("div");
        }

        public XmlElement CreateHeader1()
        {
            XmlElement result = document.CreateElement("h1");
            result.AppendChild(document.CreateTextNode("        "));
            return result;
        }

        public XmlElement CreateHeader2()
        {
            XmlElement result = document.CreateElement("h2");
            result.AppendChild(document.CreateTextNode("    "));
            return result;
        }

        public XmlElement CreateParagraph()
        {
            return document.CreateElement("p");
        }

        public XmlElement CreateTable()
        {
            return document.CreateElement("table");
        }

        public XmlElement CreateTableBody()
        {
            return document.CreateElement("tbody");
        }

        public XmlElement CreateTableCell()
        {
            return document.CreateElement("td");
        }

        public XmlElement CreateTableRow()
        {
            return document.CreateElement("tr");
        }

        public XmlText CreateText(String data)
        {
            return document.CreateTextNode(data);
        }

        public XmlElement CreateUnorderedList()
        {
            return document.CreateElement("ul");
        }

        public XmlElement Body
        {
            get
            {
                return body;
            }
        }

        public XmlDocument Document
        {
            get
            {
                return document;
            }
        }

        public XmlElement Head
        {
            get
            {
                return head;
            }
        }

        public String Title
        {
            get
            {
                if (title == null)
                    return null;

                return titleText.InnerText;
            }
            set
            {
                if (string.IsNullOrEmpty(value) && this.title != null)
                {
                    this.head.RemoveChild(this.title);
                    this.title = null;
                    this.titleText = null;
                }

                if (this.title == null)
                {
                    this.title = document.CreateElement("title");
                    this.titleText = document.CreateTextNode(value);
                    this.title.AppendChild(this.titleText);
                    this.head.AppendChild(title);
                }

                this.titleText.InnerText=(value);
            }
        }

       
    }
}
