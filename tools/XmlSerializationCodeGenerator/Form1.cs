
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

namespace XmlSerializationCodeGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var rootNode = treeView1.Nodes.Add("CT_Drawing");
            RecursiveRun(typeof(CT_Drawing), rootNode, 0);
            //treeView1.ExpandAll();
        }
        void RecursiveRun(Type c, TreeNode node, int level)
        {
            if (c.Name == "XmlElement"||c.Name=="Byte[]")
                return;
            
            node.Tag = c;
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
            foreach (var p in properties)
            {
                if (p.Name.EndsWith("Specified"))
                    continue;
                if (p.PropertyType.IsValueType)
                {
                    if (p.PropertyType.Name == "Int32")
                    {
                        sb.AppendLine(string.Format("\tctObj.{0} = XmlHelper.ReadInt(node.Attributes[\"{0}\"]);", p.Name));
                    }
                    else if (p.PropertyType.Name == "Int64")
                    {
                        sb.AppendLine(string.Format("\tctObj.{0} = XmlHelper.ReadLong(node.Attributes[\"{0}\"]);", p.Name));
                    }
                    else if (p.PropertyType.Name == "UInt32")
                    {
                        sb.AppendLine(string.Format("\tctObj.{0} = XmlHelper.ReadUInt(node.Attributes[\"{0}\"]);", p.Name));
                    }
                    else if (p.PropertyType.Name == "Boolean")
                    {
                        sb.AppendLine(string.Format("\tctObj.{0} = XmlHelper.ReadBool(node.Attributes[\"{0}\"]);", p.Name));
                    }
                    else if (p.PropertyType.IsEnum)
                    {
                        sb.AppendLine(string.Format("\tif (node.Attributes[\"{0}\"]!=null)", p.Name));
                        sb.AppendLine(string.Format("\t\tctObj.{0} = ({1})Enum.Parse(typeof({1}), node.Attributes[\"{0}\"].Value);", p.Name, p.PropertyType.Name));
                    }
                }
                else if (p.PropertyType.Name == "String")
                {
                    sb.AppendLine(string.Format("\tctObj.{0} = XmlHelper.ReadString(node.Attributes[\"{0}\"]);", p.Name));
                }
                else if (p.PropertyType.Name == "Byte[]")
                {
                    sb.AppendLine(string.Format("\tctObj.{0} = XmlHelper.ReadBytes(node.Attributes[\"{0}\"]);", p.Name));
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
                    if (p.PropertyType.GetProperties().Length == 0)
                        sb.AppendLine(string.Format("\t\t\tctObj.{0} = new {1}();", p.Name, p.PropertyType.Name));
                    else
                        sb.AppendLine(string.Format("\t\t\tctObj.{0} = {1}.Parse(childNode, namespaceManager);", p.Name, p.PropertyType.Name));
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
                        sb.AppendLine(string.Format("\t\t\tctObj.{0}.Add({1}.Parse(childNode, namespaceManager));", p.Name, genericType.Name));
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
            sb.AppendLine("internal void Write(StreamWriter sw, string nodeName)");
            sb.AppendLine("{");
            sb.AppendLine("\tsw.Write(string.Format(\"<{0}\",nodeName));");
            foreach (var p in properties)
            {
                if (p.Name.EndsWith("Specified"))
                    continue;
                if (p.PropertyType.IsValueType)
                {
                    if (p.PropertyType.IsEnum)
                    {
                        sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{0}\", this.{0}.ToString());", p.Name));
                    }
                    else
                    {
                        sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{0}\", this.{0});", p.Name));
                    }
                }
                else if (p.PropertyType.Name == "String")
                {
                    sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{0}\", this.{0});", p.Name));
                }
                else if (p.PropertyType.Name == "Byte[]")
                {
                    sb.AppendLine(string.Format("\tXmlHelper.WriteAttribute(sw, \"{0}\", this.{0});", p.Name));
                }
            }
            sb.AppendLine("\tsw.Write(\">\");");
            foreach (var p in subProperties)
            {
                sb.AppendLine(string.Format("\tif(this.{0}!=null)", p.Name));
                if (p.PropertyType.GetProperties().Length == 0)
                {
                    sb.AppendLine(string.Format("\t\tsw.Write(\"<{0}/>\");", p.Name));
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
                sb.AppendLine(string.Format("\tforeach({0} x in this.{1})", genericType.Name, p.Name));
                sb.AppendLine("\t{");
                if (genericType.IsEnum)
                {
                    sb.AppendLine("\t\tsw.Write(string.Format(\"<{0}/>\",x));");
                }
                else
                {
                    sb.AppendLine(string.Format("\t\tx.Write(sw,\"{0}\");",p.Name));
                }
                sb.AppendLine("\t}");
            }
            sb.AppendLine("\tsw.Write(string.Format(\"</{0}>\",nodeName));");
            sb.AppendLine("}");
            #endregion
            textBox1.Text= sb.ToString();
        }
    }
}
