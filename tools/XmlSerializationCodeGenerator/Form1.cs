
using ICSharpCode.TextEditor.Document;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace XmlSerializationCodeGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            editCSharp.Document.HighlightingStrategy = HighlightingStrategyFactory.CreateHighlightingStrategy("C#");
            editCSharp.Encoding = System.Text.Encoding.UTF8;
        }
        List<string> keywords = new List<string>() { 
            "in", "out", "ref", "string", "double", "int", "float", "uint", "long", "ulong", "byte"
        };
        List<Type> types = new List<Type>();

        private void button1_Click(object sender, EventArgs e)
        {

            Type targetType = typeof(NPOI.OpenXmlFormats.Spreadsheet.CT_PivotTableDefinition);
            var rootNode = treeView1.Nodes.Add(targetType.Name);
            RecursiveRun(targetType, rootNode, 0);
            //treeView1.ExpandAll();

            StringBuilder sb=new StringBuilder();
            foreach(Type type in types)
            {
                if (type.GetProperties().Length == 0)
                    sb.AppendLine(type.Name + " [x]");
                else
                {
                    sb.AppendLine(type.Name);
                    if (type.GetMethod("Parse", BindingFlags.Static | BindingFlags.Public) != null)
                        sb.AppendLine("- Parse");

                    if (type.GetMethod("Write", BindingFlags.NonPublic | BindingFlags.Instance) != null)
                        sb.AppendLine("- Write");
                }
            }
            editCSharp.Text = sb.ToString();
        }
        void RecursiveRun(Type c, TreeNode node, int level)
        {
            if (c.Name == "XmlElement"||c.Name=="Byte[]")
                return;

            node.Tag = c;

            if (level >7)
                return;

            var properties = c.GetProperties();
            foreach (var p in properties)
            {
                if (p.Name.EndsWith("Specified")||p.Name=="Item")
                    continue;

                if (p.PropertyType.IsClass&& !(p.PropertyType==typeof(string)))
                {
                    var subNode = node.Nodes.Add(p.Name + "["+p.PropertyType.Name + " class]");
                    if (typeof(IList).IsAssignableFrom(p.PropertyType)
                        && p.PropertyType.IsGenericType)
                    {
                        Type genericType = p.PropertyType.GetGenericArguments()[0];
                        if (!types.Contains(p.PropertyType.GetGenericArguments()[0]))
                            types.Add(genericType);

                        subNode.Text= subNode.Text.Replace("`1", "<" + genericType.Name+">");
                        RecursiveRun(p.PropertyType.GetGenericArguments()[0], subNode, level + 1);
                        //textBox1.Text += c.Name + " - " + p.PropertyType.GetGenericArguments()[0].Name + Environment.NewLine;
                    }
                    else if (p.PropertyType.BaseType!=null&&p.PropertyType.BaseType.Name == "Array")
                    {
                        //textBox1.Text += c.Name +" - "+p.PropertyType+ Environment.NewLine;
                        RecursiveRun(p.PropertyType, subNode, level + 1);
                    }
                    else
                    {
                        if (!types.Contains(p.PropertyType))
                            types.Add(p.PropertyType);

                        RecursiveRun(p.PropertyType, subNode, level + 1);
                    }
                }
                else if (p.PropertyType.IsValueType)
                {
                    node.Nodes.Add(p.Name + "[" + p.PropertyType.Name + " property]");
                }
                else
                {
                    node.Nodes.Add(p.Name + "[" + p.PropertyType.Name + " property]");
                }
            }            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode == null)
                return;

            Type t = (Type)treeView1.SelectedNode.Tag;

            #region generate parse code
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("public static {0} Parse(XmlNode node, XmlNamespaceManager namespaceManager)"+Environment.NewLine, t.Name);
            sb.AppendLine("{");
            sb.AppendLine("\tif(node==null)");
            sb.AppendLine("\t\treturn null;");
            sb.AppendLine(string.Format("\t{0} ctObj = new {0}();", t.Name));

            List<PropertyInfo> subProperties = new List<PropertyInfo>();
            List<PropertyInfo> listProps = new List<PropertyInfo>();

            var properties = t.GetProperties();
            string varName;
            foreach (var p in properties)
            {
                //忽略Specified结尾的Property
                if (p.Name.EndsWith("Specified"))
                    continue;
                varName = p.Name;
                if (this.keywords.Contains(p.Name))
                    varName = "@" + p.Name;
                if (p.GetCustomAttributes(typeof(XmlElementAttribute), false).Length > 0)
                {
                    if (typeof(IList).IsAssignableFrom(p.PropertyType)
                        && p.PropertyType.IsGenericType)
                        listProps.Add(p);
                    else
                        subProperties.Add(p);
                    continue;
                }
                string attributePrefix = GetXmlAttributePrefix(p);
                if (p.PropertyType.IsValueType)
                {
                    sb.AppendLine(string.Format("\tif (node.Attributes[\"{1}{0}\"]!=null)", p.Name, attributePrefix));
                    sb.Append("\t");

                    if (p.PropertyType.Name == "Int32")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadInt(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "Int64")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadLong(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "Double")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadDouble(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "Byte")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadByte(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "SByte")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadSByte(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "Int16")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadShort(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "UInt16")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadUShort(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "UInt32")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadUInt(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "UInt64")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadULong(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "Boolean")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadBool(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "Decimal")
                    {
                        sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadDecimal(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                    }
                    else if (p.PropertyType.Name == "DateTime")
                    {
                        sb.AppendFormat("\tctObj.{2} = XmlHelper.ReadDateTime(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName).AppendLine();
                    }
                    else if (p.PropertyType.IsEnum)
                    {
                        if (p.PropertyType.Name == "ST_TrueFalse")
                            sb.AppendLine(string.Format("\tctObj.{0} = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalse(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix));
                        else if (p.PropertyType.Name == "ST_TrueFalseBlank")
                            sb.AppendLine(string.Format("\tctObj.{0} = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalseBlank(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix));
                        else
                            sb.AppendLine(string.Format("\t\tctObj.{0} = ({1})Enum.Parse(typeof({1}), node.Attributes[\"{2}{0}\"].Value);", p.Name, p.PropertyType.Name, attributePrefix));
                    }
                }
                else if (p.PropertyType.Name == "String")
                {
                    sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadString(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                }
                else if (p.PropertyType.Name == "Byte[]")
                {
                    sb.AppendLine(string.Format("\tctObj.{2} = XmlHelper.ReadBytes(node.Attributes[\"{1}{0}\"]);", p.Name, attributePrefix, varName));
                }
                else if (p.PropertyType.IsClass)
                {
                    if (typeof(IList).IsAssignableFrom(p.PropertyType)
                        && p.PropertyType.IsGenericType)
                    {
                        listProps.Add(p);
                    }
                    else
                    {
                        subProperties.Add(p);
                        //sb.AppendLine(string.Format("\tctObj.{0} = {1}.Parse(node, namespaceManager);", p.Name, p.PropertyType.Name));
                    }
                }
            }
            foreach (var p in listProps)
            {
                sb.AppendLine(string.Format("\tctObj.{0}=new List<{1}>();", p.Name, p.PropertyType.GetGenericArguments()[0].Name));
            }
            if (listProps.Count > 0 || subProperties.Count > 0)
            {
                sb.AppendLine("\tforeach(XmlNode childNode in node.ChildNodes)");
                sb.AppendLine("\t{");
                bool firstIf = true;
                foreach (var p in subProperties)
                {
                    if (firstIf)
                    {
                        sb.AppendLine(string.Format("\t\tif(childNode.LocalName == \"{0}\")", p.Name));
                        firstIf = false;
                    }
                    else
                    {
                        sb.AppendLine(string.Format("\t\telse if(childNode.LocalName == \"{0}\")", p.Name));
                    }
                    if (p.PropertyType.IsValueType)
                    {
                        sb.AppendLine(string.Format("\t\t\tctObj.{0} = {1}.Parse(childNode.InnerText);", p.Name, p.PropertyType.Name));
                    }
                    else if (p.PropertyType.Name == "String")
                    {
                        sb.AppendLine(string.Format("\t\t\tctObj.{0} = childNode.InnerText;", p.Name));
                    }
                    else if (p.PropertyType.GetProperties().Length == 0)
                    {
                        sb.AppendLine(string.Format("\t\t\tctObj.{0} = new {1}();", p.Name, p.PropertyType.Name));
                    }
                    else
                    {
                        sb.AppendLine(string.Format("\t\t\tctObj.{0} = {1}.Parse(childNode, namespaceManager);", p.Name, p.PropertyType.Name));
                    }
                }
                foreach (var p in listProps)
                {
                    Type genericType = p.PropertyType.GetGenericArguments()[0];
                    if (genericType.IsEnum)
                    {
                        foreach (var enumName in genericType.GetEnumNames())
                        {
                            if (firstIf)
                            {
                                sb.AppendLine(string.Format("\t\tif(childNode.LocalName == \"{0}\")", enumName));
                                firstIf = false;
                            }
                            else
                            {
                                sb.AppendLine(string.Format("\t\telse if(childNode.LocalName == \"{0}\")", enumName));
                            }
                            sb.AppendLine(string.Format("\t\t\tctObj.{0}.Add({1}.{2});", p.Name, genericType.Name, enumName));
                        }
                    }
                    else
                    {
                        if (firstIf)
                        {
                            sb.AppendLine(string.Format("\t\tif(childNode.LocalName == \"{0}\")", p.Name));
                            firstIf = false;
                        }
                        else
                        {
                            sb.AppendLine(string.Format("\t\telse if(childNode.LocalName == \"{0}\")", p.Name));
                        }
                        if (genericType.Name == "String")
                        {
                            sb.AppendLine(string.Format("\t\t\tctObj.{0}.Add(childNode.InnerText);", p.Name, genericType.Name));
                        }
                        else if (genericType.GetProperties().Length == 0)
                        {
                            sb.AppendLine(string.Format("\t\t\tctObj.{0}.Add(new {1}());", p.Name, genericType.Name));
                        }
                        else
                        {
                            sb.AppendLine(string.Format("\t\t\tctObj.{0}.Add({1}.Parse(childNode, namespaceManager));", p.Name, genericType.Name));
                        }
                    }
                }
                sb.AppendLine("\t}");
            }
            sb.AppendLine("\treturn ctObj;");
            sb.AppendLine("}");
            #endregion
            #region generate write code
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine();
            string xmlPrefix = GetXmlPrefix(t);
            sb.AppendLine("internal void Write(StreamWriter sw, string nodeName)");
            sb.AppendLine("{");
            sb.AppendLine("\tsw.Write(string.Format(\"<"+xmlPrefix+"{0}\",nodeName));");
            foreach (var p in properties)
            {
                if (p.Name.EndsWith("Specified"))
                    continue;
                if (p.GetCustomAttributes(typeof(XmlElementAttribute), false).Length > 0)
                {
                    continue;
                }
                varName = p.Name;
                if (this.keywords.Contains(p.Name))
                    varName = "@" + p.Name;
                string attributePrefix = GetXmlAttributePrefix(p);
                if (p.PropertyType.IsValueType)
                {
                    if (p.PropertyType.IsEnum)
                    {
                        if (p.PropertyType.Name == "ST_TrueFalse" || p.PropertyType.Name == "ST_TrueFalseBlank")
                            sb.AppendLine(string.Format("\tNPOI.OpenXmlFormats.Util.XmlHelper.WriteAttribute(sw, \"{1}{0}\", this.{2});", p.Name, attributePrefix, varName));
                        else if (p.PropertyType.Name == "ST_Ext")
                        {
                            sb.AppendLine("\tif(this.ext!=ST_Ext.NONE)");
                            sb.AppendLine(string.Format("\t\tXmlHelper.WriteAttribute(sw, \"{1}{0}\", this.{2}.ToString());", p.Name, attributePrefix, varName));
                        }
                        else
                            sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{1}{0}\", this.{2}.ToString());", p.Name, attributePrefix, varName));
                    }
                    else
                    {
                        sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{1}{0}\", this.{2});", p.Name, attributePrefix, varName));
                    }
                }
                else if (p.PropertyType.Name == "String")
                {
                    sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{1}{0}\", this.{2});", p.Name, attributePrefix, varName));
                }
                else if (p.PropertyType.Name == "Byte[]")
                {
                    sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{1}{0}\", this.{2});", p.Name, attributePrefix, varName));
                }
            }
            sb.AppendLine("\tsw.Write(\">\");");
            foreach (var p in subProperties)
            {
                varName = p.Name;
                if (this.keywords.Contains(p.Name))
                    varName = "@" + p.Name;
                sb.AppendLine(string.Format("\tif(this.{0}!=null)", varName));
                if (p.PropertyType.IsValueType)
                {
                    sb.AppendLine(string.Format("\t\tsw.Write(string.Format(\"<{1}{0}>{{0}}</{1}{0}>\",this.{2}));", p.Name, GetXmlPrefix(p.PropertyType), varName));
                }
                else if (p.PropertyType.Name == "String")
                {
                    sb.AppendLine(string.Format("\t\tsw.Write(string.Format(\"<{1}{0}>{{0}}</{1}{0}>\",this.{2}));", p.Name, GetXmlPrefix(p.PropertyType), varName));
                }
                else if (p.PropertyType.Name == "Boolean")
                {
                    var defaultAttrs = p.GetCustomAttributes(typeof(DefaultValueAttribute), false);
                    if (defaultAttrs.Length > 0)
                    {
                        if (((bool)((DefaultValueAttribute)defaultAttrs[0]).Value) == true)
                        {
                            sb.AppendLine(string.Format("\t\tsw.Write(string.Format(\"<{1}{0}>{{0}}</{1}{0}>\",this.{2}, true));", p.Name, GetXmlPrefix(p.PropertyType), varName));
                        }
                        else
                        {
                            sb.AppendLine(string.Format("\t\tsw.Write(string.Format(\"<{1}{0}>{{0}}</{1}{0}>\",this.{2}));", p.Name, GetXmlPrefix(p.PropertyType), varName));
                        }
                    }
                    else
                    {
                        sb.AppendLine(string.Format("\t\tsw.Write(string.Format(\"<{1}{0}>{{0}}</{1}{0}>\",this.{2}));", p.Name, GetXmlPrefix(p.PropertyType), varName));
                    }
                }
                else if (p.PropertyType.GetProperties().Length == 0)
                {
                    sb.AppendLine(string.Format("\t\tsw.Write(\"<{1}{0}/>\");", p.Name, GetXmlPrefix(p.PropertyType)));
                }
                else
                {
                    sb.AppendLine(string.Format("\t\tthis.{0}.Write(sw, \"{0}\");", p.Name));
                }
            }
            foreach (var p in listProps)
            {
                Type genericType = p.PropertyType.GetGenericArguments()[0];
                if (genericType.Name == "Object")
                    continue;
                sb.AppendLine(string.Format("\tif(this.{0}!=null)", p.Name));
                sb.AppendLine("\t{");
                sb.AppendLine(string.Format("\t\tforeach({0} x in this.{1})", genericType.Name, p.Name));
                sb.AppendLine("\t\t{");
                if (genericType.IsEnum)
                {
                    sb.AppendLine("\t\t\tsw.Write(string.Format(\"<{0}/>\",x));");
                }
                else if (genericType.Name == "String")
                {
                    sb.AppendLine(string.Format("\t\t\tsw.Write(string.Format(\"<{0}>{{0}}</{0}>\",x));", p.Name));
                }
                else if (genericType.GetProperties().Length == 0)
                {
                    sb.AppendLine(string.Format("\t\tsw.Write(\"<{0}/>\");", p.Name));
                }
                else
                {
                    sb.AppendLine(string.Format("\t\tx.Write(sw,\"{0}\");", p.Name));
                }
                sb.AppendLine("\t\t}");
                sb.AppendLine("\t}");
            }
            sb.AppendLine("\tsw.Write(string.Format(\"</" + xmlPrefix + "{0}>\",nodeName));");
            sb.AppendLine("}");
            #endregion
            editCSharp.Text = sb.ToString();

            Clipboard.SetText(editCSharp.Text);
        }
        public string GetXmlAttributePrefix(PropertyInfo p)
        { 
            var a = p.GetCustomAttributes(typeof(XmlAttributeAttribute), false);
            
            if (a.Length == 0)
                return "";
            string n = ((XmlAttributeAttribute)a[0]).Namespace;
            if (n == "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
            {
                return "xdr:";
            }
            else if (n == "http://schemas.openxmlformats.org/drawingml/2006/main")
            {
                return "a:";
            }
            else if (n == "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            {
                return "";
            }
            else if (n == "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            {
                return "r:";
            }
            else if (n == "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            {
                return "w:";
            }
            else if (n == "http://schemas.openxmlformats.org/markup-compatibility/2006")
            {
                return "ve:";
            }
            else if (n == "urn:schemas-microsoft-com:office:office")
            {
                return "o:";
            }
            else if (n == "urn:schemas-microsoft-com:office:word")
            {
                return "w:";
            }
            else if (n == "urn:schemas-microsoft-com:vml")
            {
                return "v:";
            }
            else if (n == "urn:schemas-microsoft-com:office:excel")
            {
                return "x:";
            }
            else if (n == "urn:schemas-microsoft-com:office:powerpoint")
            {
                return "p:";
            }
            else if (n == "http://schemas.openxmlformats.org/officeDocument/2006/math")
            {
                return "m:";
            }
            else if (n == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
            {
                return "wp:";
            }
            else if (n == "http://schemas.openxmlformats.org/drawingml/2006/chart")
            {
                return "c:";
            }
            return ""; 
        }
        public string GetXmlPrefix(Type p)
        {
            var a = p.GetCustomAttributes(typeof(XmlTypeAttribute), false);
            if (a.Length == 0)
                return "";
            string n = ((XmlTypeAttribute)a[0]).Namespace;
            if (n == "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
            {
                return "xdr:";
            }
            else if (n == "http://schemas.openxmlformats.org/drawingml/2006/main")
            {
                return "a:";
            }
            else if (n == "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            {
                return "";
            }
            else if (n == "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            {
                return "r:";
            }
            else if (n == "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            {
                return "w:";
            }
            else if (n == "http://schemas.openxmlformats.org/markup-compatibility/2006")
            {
                return "ve:";
            }
            else if (n == "urn:schemas-microsoft-com:office:office")
            {
                return "o:";
            }
            else if (n == "urn:schemas-microsoft-com:vml")
            {
                return "v:";
            }
            else if (n == "urn:schemas-microsoft-com:office:excel")
            {
                return "x:";
            }
            else if (n == "urn:schemas-microsoft-com:office:powerpoint")
            {
                return "p:";
            }
            else if (n == "urn:schemas-microsoft-com:office:word")
            {
                return "w:";
            }
            else if (n == "http://schemas.openxmlformats.org/officeDocument/2006/math")
            {
                return "m:";
            }
            else if (n == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
            {
                return "wp:";
            }
            else if (n == "http://schemas.openxmlformats.org/drawingml/2006/chart")
            {
                return "c:";
            }
            return "";    
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(editCSharp.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            StringBuilder sb=new StringBuilder();
            Type type = (Type)treeView1.SelectedNode.Tag;
            var property = type.GetProperty("Items");

            if (property != null)
            {
                var attrs = property.GetCustomAttributes(typeof(XmlElementAttribute), false);
                foreach (var attr in attrs)
                {
                    var xmlAttr = (XmlElementAttribute)attr;
                    sb.AppendLine(string.Format("List<{1}> {0}Field;", xmlAttr.ElementName, xmlAttr.Type.Name));
                    sb.AppendLine(string.Format("public List<{1}> {0}", xmlAttr.ElementName, xmlAttr.Type.Name));
                    sb.AppendLine("{");
                    sb.AppendLine(string.Format("\tget{{return this.{0}Field;}}", xmlAttr.ElementName));
                    sb.AppendLine(string.Format("\tset{{this.{0}Field=value;}}", xmlAttr.ElementName));
                    sb.AppendLine("}");
                    sb.AppendLine();
                }
                editCSharp.Text = sb.ToString();
            }
            else
            {
               property = type.GetProperty("Item");
               if (property != null)
               {
                   var attrs = property.GetCustomAttributes(typeof(XmlElementAttribute), false);
                   foreach (var attr in attrs)
                   {
                       var xmlAttr = (XmlElementAttribute)attr;
                       sb.AppendLine(string.Format("{1} {0}Field;", xmlAttr.ElementName, xmlAttr.Type.Name));
                       sb.AppendLine(string.Format("public {1} {0}", xmlAttr.ElementName, xmlAttr.Type.Name));
                       sb.AppendLine("{");
                       sb.AppendLine(string.Format("\tget{{return this.{0}Field;}}", xmlAttr.ElementName));
                       sb.AppendLine(string.Format("\tset{{this.{0}Field=value;}}", xmlAttr.ElementName));
                       sb.AppendLine("}");
                       sb.AppendLine();
                   }
                   editCSharp.Text = sb.ToString();
               }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode == null)
                return;

            StringBuilder writeCode = new StringBuilder();
            StringBuilder parseCode = new StringBuilder();
            Type type = (Type)treeView1.SelectedNode.Tag;
            string xmlPrefix = GetXmlPrefix(type);

            parseCode.AppendFormat("public static {0} Parse(XmlNode node, XmlNamespaceManager namespaceManager)" + Environment.NewLine, type.Name);
            parseCode.AppendLine("{");
            parseCode.AppendLine("\tif(node==null)");
            parseCode.AppendLine("\t\treturn null;");
            parseCode.AppendLine(string.Format("\t{0} ctObj = new {0}();", type.Name));
            parseCode.AppendLine("\tforeach(XmlNode childNode in node.ChildNodes)");
            parseCode.AppendLine("\t{");

            writeCode.AppendLine("internal void Write(StreamWriter sw, string nodeName)");
            writeCode.AppendLine("{");
            writeCode.AppendLine("\tsw.Write(string.Format(\"<" + xmlPrefix + "{0}\",nodeName));");
            writeCode.AppendLine("\tsw.Write(\">\");");
            writeCode.AppendLine("\tforeach(object o in this.Items)");
            writeCode.AppendLine("\t{");
            var property = type.GetProperty("Items");

            if (property != null)
            {
                var attrs = property.GetCustomAttributes(typeof(XmlElementAttribute), false);

                var firstIf = true;
                foreach (var attr in attrs)
                {
                    var xmlAttr = (XmlElementAttribute)attr;

                    if (firstIf)
                    {
                        parseCode.AppendLine(string.Format("\t\tif(childNode.LocalName == \"{0}\")", xmlAttr.ElementName));
                    }
                    else
                    {
                        parseCode.AppendLine(string.Format("\t\telse if(childNode.LocalName == \"{0}\")", xmlAttr.ElementName));
                    }
                    parseCode.AppendLine("\t\t{");
                    if (xmlAttr.Type.GetProperties().Length == 0)
                    {
                        parseCode.AppendLine(string.Format("\t\t\tctObj.Items.Add(new {0}());", xmlAttr.Type.Name));
                    }
                    else
                    {
                        parseCode.AppendLine(string.Format("\t\t\tctObj.Items.Add({0}.Parse(childNode, namespaceManager));", xmlAttr.Type.Name));
                    }
                    parseCode.AppendLine(string.Format("\t\t\tctObj.ItemsElementName.Add(ItemsChoiceType{1}.{0});", xmlAttr.ElementName, textBox2.Text));                    
                    parseCode.AppendLine("\t\t}");
                    if (firstIf)
                    {
                        writeCode.AppendLine(string.Format("\t\tif(o is {0})", xmlAttr.Type.Name));
                    }
                    else
                    {
                        writeCode.AppendLine(string.Format("\t\telse if(o is {0})", xmlAttr.Type.Name));
                    }
                    if(xmlAttr.Type.GetProperties().Length==0)
                        writeCode.AppendLine(string.Format("\t\t\tsw.Write(\"<{0}/>\");", xmlAttr.ElementName));
                    else
                        writeCode.AppendLine(string.Format("\t\t\t(({0})o).Write(sw, \"{1}\");", xmlAttr.Type.Name, xmlAttr.ElementName));
                    firstIf = false;
                }
            }
            parseCode.AppendLine("\t}");
            parseCode.AppendLine("\treturn ctObj;"); 
            parseCode.AppendLine("}");

            writeCode.AppendLine("\t}");
            writeCode.AppendLine("\tsw.Write(string.Format(\"</" + xmlPrefix + "{0}>\",nodeName));");
            writeCode.AppendLine("}");

            editCSharp.Text = parseCode.ToString();
            editCSharp.Text += Environment.NewLine;
            editCSharp.Text += writeCode.ToString();
        }
    }
}
