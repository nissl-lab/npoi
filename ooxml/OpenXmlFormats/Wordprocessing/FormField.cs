using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFTextType
    {

        private ST_FFTextType valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FFTextType val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        public static CT_FFTextType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFTextType ctObj = new CT_FFTextType();
            ctObj.valField = (ST_FFTextType)Enum.Parse(typeof(ST_FFTextType), XmlHelper.ReadString(node.Attributes["w:val"]));
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.valField.ToString());
            sw.Write("/>");
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_FFTextType
    {

    
        regular,

    
        number,

    
        date,

    
        currentTime,

    
        currentDate,

    
        calculated,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFName
    {

        private string valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        public static CT_FFName Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFName ctObj = new CT_FFName();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val, true);
            sw.Write("/>");
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FldChar
    {

        private object itemField;

        private ST_FldCharType fldCharTypeField;

        private ST_OnOff fldLockField;

        private bool fldLockFieldSpecified;

        private ST_OnOff dirtyField;

        private bool dirtyFieldSpecified;

        private CT_FFData ffDataField;
        public CT_FFData ffData
        {
            get { return this.ffDataField; }
            set { this.ffDataField = value; }
        }
        private CT_Text fldDataField;
        public CT_Text fldData
        {
            get { return this.fldDataField; }
            set { this.fldDataField = value; }
        }
        private CT_TrackChangeNumbering numberingChangeField;
        public CT_TrackChangeNumbering numberingChange
        {
            get { return this.numberingChangeField; }
            set { this.numberingChangeField = value; }
        }
        [XmlElement("ffData", typeof(CT_FFData), Order = 0)]
        [XmlElement("fldData", typeof(CT_Text), Order = 0)]
        [XmlElement("numberingChange", typeof(CT_TrackChangeNumbering), Order = 0)]
        public object Item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }
        public static CT_FldChar Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FldChar ctObj = new CT_FldChar();
            if (node.Attributes["w:fldCharType"] != null)
                ctObj.fldCharType = (ST_FldCharType)Enum.Parse(typeof(ST_FldCharType), node.Attributes["w:fldCharType"].Value);
            if (node.Attributes["w:fldLock"] != null)
                ctObj.fldLock = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:fldLock"].Value);
            if (node.Attributes["w:dirty"] != null)
                ctObj.dirty = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:dirty"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "ffData")
                {
                    ctObj.ffDataField = CT_FFData.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "fldData")
                {
                    ctObj.fldDataField = CT_Text.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "numberingChange")
                {
                    ctObj.numberingChangeField = CT_TrackChangeNumbering.Parse(childNode, namespaceManager);
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:fldCharType", this.fldCharType.ToString());
            XmlHelper.WriteAttribute(sw, "w:fldLock", this.fldLock.ToString());
            XmlHelper.WriteAttribute(sw, "w:dirty", this.dirty.ToString());
            sw.Write(">");

            if (this.ffDataField != null)
            {
                this.ffDataField.Write(sw, "ffData");
            }
            if (this.fldDataField != null)
                this.fldDataField.Write(sw, "fldData");
            if (this.numberingChangeField != null)
            {
                this.numberingChangeField.Write(sw, "numberingChange");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FldCharType fldCharType
        {
            get
            {
                return this.fldCharTypeField;
            }
            set
            {
                this.fldCharTypeField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff fldLock
        {
            get
            {
                return this.fldLockField;
            }
            set
            {
                this.fldLockField = value;
            }
        }

        [XmlIgnore]
        public bool fldLockSpecified
        {
            get
            {
                return this.fldLockFieldSpecified;
            }
            set
            {
                this.fldLockFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff dirty
        {
            get
            {
                return this.dirtyField;
            }
            set
            {
                this.dirtyField = value;
            }
        }

        [XmlIgnore]
        public bool dirtySpecified
        {
            get
            {
                return this.dirtyFieldSpecified;
            }
            set
            {
                this.dirtyFieldSpecified = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFData
    {

        private List<object> itemsField;

        private List<FFDataItemsType> itemsElementNameField;

        public CT_FFData()
        {
            this.itemsElementNameField = new List<FFDataItemsType>();
            this.itemsField = new List<object>();
        }

        [XmlElement("calcOnExit", typeof(CT_OnOff), Order = 0)]
        [XmlElement("checkBox", typeof(CT_FFCheckBox), Order = 0)]
        [XmlElement("ddList", typeof(CT_FFDDList), Order = 0)]
        [XmlElement("enabled", typeof(CT_OnOff), Order = 0)]
        [XmlElement("entryMacro", typeof(CT_MacroName), Order = 0)]
        [XmlElement("exitMacro", typeof(CT_MacroName), Order = 0)]
        [XmlElement("helpText", typeof(CT_FFHelpText), Order = 0)]
        [XmlElement("name", typeof(CT_FFName), Order = 0)]
        [XmlElement("statusText", typeof(CT_FFStatusText), Order = 0)]
        [XmlElement("textInput", typeof(CT_FFTextInput), Order = 0)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                this.itemsField.Clear();
                this.itemsField.AddRange(value);
            }
        }

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public FFDataItemsType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                this.itemsElementNameField.Clear();
                this.itemsElementNameField.AddRange(value);
            }
        }

        internal static CT_FFData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFData ctObj = new CT_FFData();

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "name")
                {
                    ctObj.AddNewObject(CT_FFName.Parse(childNode, namespaceManager) ,FFDataItemsType.name);
                }
                //else if (childNode.LocalName == "tabIndex")
                //{

                //}
                else if (childNode.LocalName == "enabled")
                {
                    ctObj.AddNewObject(CT_OnOff.Parse(childNode, namespaceManager), FFDataItemsType.name);
                }
                else if (childNode.LocalName == "calcOnExit")
                {
                    ctObj.AddNewObject(CT_OnOff.Parse(childNode, namespaceManager), FFDataItemsType.calcOnExit);
                }
                else if (childNode.LocalName == "checkBox")
                {
                    ctObj.AddNewObject(CT_FFCheckBox.Parse(childNode, namespaceManager), FFDataItemsType.checkBox);
                }
                else if (childNode.LocalName == "ddList")
                {
                    ctObj.AddNewObject(CT_FFDDList.Parse(childNode, namespaceManager), FFDataItemsType.ddList);
                }
                else if (childNode.LocalName == "entryMacro")
                {
                    ctObj.AddNewObject(CT_MacroName.Parse(childNode, namespaceManager), FFDataItemsType.entryMacro);
                }
                else if (childNode.LocalName == "exitMacro")
                {
                    ctObj.AddNewObject(CT_MacroName.Parse(childNode, namespaceManager), FFDataItemsType.exitMacro);
                }
                else if (childNode.LocalName == "helpText")
                {
                    ctObj.AddNewObject(CT_FFHelpText.Parse(childNode, namespaceManager), FFDataItemsType.helpText);
                }
                else if (childNode.LocalName == "statusText")
                {
                    ctObj.AddNewObject(CT_FFStatusText.Parse(childNode, namespaceManager), FFDataItemsType.statusText);
                }
                else if (childNode.LocalName == "textInput")
                {
                    ctObj.AddNewObject(CT_FFTextInput.Parse(childNode, namespaceManager), FFDataItemsType.textInput);
                }
            }
            return ctObj;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));

            for (int i=0;i<this.itemsElementNameField.Count;i++)
            {
                if (this.itemsElementNameField[i] == FFDataItemsType.name)
                    (this.itemsField[i] as CT_FFName).Write(sw, "name");
                else if (this.itemsElementNameField[i] == FFDataItemsType.enabled)
                    (this.itemsField[i] as CT_OnOff).Write(sw, "enabled");
                else if (this.itemsElementNameField[i] == FFDataItemsType.calcOnExit)
                    (this.itemsField[i] as CT_OnOff).Write(sw, "calcOnExit");
                else if (this.itemsElementNameField[i] == FFDataItemsType.ddList)
                    (this.itemsField[i] as CT_FFDDList).Write(sw, "ddList");
                else if (this.itemsElementNameField[i] == FFDataItemsType.checkBox)
                    (this.itemsField[i] as CT_FFCheckBox).Write(sw, "checkBox");
                else if (this.itemsElementNameField[i] == FFDataItemsType.entryMacro)
                    (this.itemsField[i] as CT_MacroName).Write(sw, "entryMacro");
                else if (this.itemsElementNameField[i] == FFDataItemsType.exitMacro)
                    (this.itemsField[i] as CT_MacroName).Write(sw, "exitMacro");
                else if (this.itemsElementNameField[i] == FFDataItemsType.helpText)
                    (this.itemsField[i] as CT_FFHelpText).Write(sw, "helpText");
                else if (this.itemsElementNameField[i] == FFDataItemsType.statusText)
                    (this.itemsField[i] as CT_FFStatusText).Write(sw, "statusText");
                else if (this.itemsElementNameField[i] == FFDataItemsType.textInput)
                    (this.itemsField[i] as CT_FFTextInput).Write(sw, "textInput");
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }
        private void AddNewObject(object obj, FFDataItemsType type)
        {
            lock(this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(obj);
            }
        }

        #region Generic methods for object operation

        private List<T> GetObjectList<T>(FFDataItemsType type) where T : class
        {
            lock (this)
            {
                List<T> list = new List<T>();
                for (int i = 0; i < itemsElementNameField.Count; i++)
                {
                    if (itemsElementNameField[i] == type)
                        list.Add(itemsField[i] as T);
                }
                return list;
            }
        }
        private int SizeOfObjectArray(FFDataItemsType type)
        {
            lock (this)
            {
                int size = 0;
                for (int i = 0; i < itemsElementNameField.Count; i++)
                {
                    if (itemsElementNameField[i] == type)
                        size++;
                }
                return size;
            }
        }
        private T GetObjectArray<T>(int p, FFDataItemsType type) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return null;
                return itemsField[pos] as T;
            }
        }
        private T InsertNewObject<T>(FFDataItemsType type, int p) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                this.itemsElementNameField.Insert(pos, type);
                this.itemsField.Insert(pos, t);
            }
            return t;
        }
        private T AddNewObject<T>(FFDataItemsType type) where T : class, new()
        {
            T t = new T();
            lock (this)
            {
                this.itemsElementNameField.Add(type);
                this.itemsField.Add(t);
            }
            return t;
        }
        private void SetObjectArray<T>(FFDataItemsType type, int p, T obj) where T : class
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return;
                if (this.itemsField[pos] is T)
                    this.itemsField[pos] = obj;
                else
                    throw new Exception(string.Format(@"object types are difference, itemsField[{0}] is {1}, and parameter obj is {2}",
                        pos, this.itemsField[pos].GetType().Name, typeof(T).Name));
            }
        }
        private int GetObjectIndex(FFDataItemsType type, int p)
        {
            int index = -1;
            int pos = 0;
            for (int i = 0; i < itemsElementNameField.Count; i++)
            {
                if (itemsElementNameField[i] == type)
                {
                    if (pos == p)
                    {
                        //return itemsField[p] as T;
                        index = i;
                        break;
                    }
                    else
                        pos++;
                }
            }
            return index;
        }
        private void RemoveObject(FFDataItemsType type, int p)
        {
            lock (this)
            {
                int pos = GetObjectIndex(type, p);
                if (pos < 0 || pos >= this.itemsField.Count)
                    return;
                itemsElementNameField.RemoveAt(pos);
                itemsField.RemoveAt(pos);
            }
        }
        #endregion

        public List<CT_FFCheckBox> GetCheckBoxList()
        {
            return GetObjectList<CT_FFCheckBox>(FFDataItemsType.checkBox);
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFCheckBox
    {

        private object itemField;

        private CT_OnOff defaultField;

        private CT_OnOff checkedField;

        public CT_FFCheckBox()
        {
            this.checkedField = new CT_OnOff();
            this.defaultField = new CT_OnOff();
        }

        [XmlElement("size", typeof(CT_HpsMeasure), Order = 0)]
        [XmlElement("sizeAuto", typeof(CT_OnOff), Order = 0)]
        public object Item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_OnOff @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_OnOff @checked
        {
            get
            {
                return this.checkedField;
            }
            set
            {
                this.checkedField = value;
            }
        }

        public static CT_FFCheckBox Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFCheckBox ctObj = new CT_FFCheckBox();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "checked")
                {
                    ctObj.checkedField = CT_OnOff.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "default")
                {
                    ctObj.defaultField = CT_OnOff.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "size")
                {
                    ctObj.itemField = CT_HpsMeasure.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "sizeAuto")
                {
                    ctObj.itemField = CT_OnOff.Parse(childNode, namespaceManager);
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}>", nodeName));
            if (this.defaultField != null)
                this.defaultField.Write(sw, "w:default");
            if (this.checkedField != null)
                this.checkedField.Write(sw, "w:checked");
            if (this.itemField != null)
            {
                if (this.itemField is CT_OnOff)
                    (this.itemField as CT_OnOff).Write(sw, "w:sizeAuto");
                else
                    (this.itemField as CT_HpsMeasure).Write(sw, "w:size");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFDDList
    {

        private CT_DecimalNumber resultField;

        private CT_DecimalNumber defaultField;

        private List<CT_String> listEntryField;

        public CT_FFDDList()
        {
            this.listEntryField = new List<CT_String>();
            this.defaultField = new CT_DecimalNumber();
            this.resultField = new CT_DecimalNumber();
        }

        [XmlElement(Order = 0)]
        public CT_DecimalNumber result
        {
            get
            {
                return this.resultField;
            }
            set
            {
                this.resultField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_DecimalNumber @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [XmlElement("listEntry", Order = 2)]
        public List<CT_String> listEntry
        {
            get
            {
                return this.listEntryField;
            }
            set
            {
                this.listEntryField = value;
            }
        }

        public static CT_FFDDList Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFDDList ctObj = new CT_FFDDList();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "result")
                {
                    ctObj.resultField = CT_DecimalNumber.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "default")
                {
                    ctObj.defaultField = CT_DecimalNumber.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "listEntry")
                {
                    ctObj.listEntryField.Add(CT_String.Parse(childNode, namespaceManager));
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}>", nodeName));
            if (this.defaultField != null)
                this.defaultField.Write(sw, "w:default");
            if (this.resultField != null)
                this.resultField.Write(sw, "w:result");
            foreach (CT_String str in listEntry)
            {
                str.Write(sw, "w:listEntry");
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFHelpText
    {

        private ST_InfoTextType typeField;

        private bool typeFieldSpecified;

        private string valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_InfoTextType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlIgnore]
        public bool typeSpecified
        {
            get
            {
                return this.typeFieldSpecified;
            }
            set
            {
                this.typeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        public static CT_FFHelpText Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFHelpText ctObj = new CT_FFHelpText();
            if (node.Attributes["w:type"] != null)
            {
                ctObj.typeFieldSpecified = true;
                ctObj.typeField = (ST_InfoTextType)Enum.Parse(typeof(ST_InfoTextType), node.Attributes["w:type"].Value);
            }
            ctObj.valField = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.valField);
            if (this.typeFieldSpecified)
            {
                XmlHelper.WriteAttribute(sw, "w:type", this.typeField.ToString());
            }
            sw.Write("/>");
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_InfoTextType
    {

    
        text,

    
        autoText,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFStatusText
    {

        private ST_InfoTextType typeField;

        private bool typeFieldSpecified;

        private string valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_InfoTextType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlIgnore]
        public bool typeSpecified
        {
            get
            {
                return this.typeFieldSpecified;
            }
            set
            {
                this.typeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        public static CT_FFStatusText Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFStatusText ctObj = new CT_FFStatusText();
            if (node.Attributes["w:type"] != null)
            {
                ctObj.typeFieldSpecified = true;
                ctObj.typeField = (ST_InfoTextType)Enum.Parse(typeof(ST_InfoTextType), node.Attributes["w:type"].Value);
            }
            ctObj.valField = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.valField);
            if (this.typeFieldSpecified)
            {
                XmlHelper.WriteAttribute(sw, "w:type", this.typeField.ToString());
            }
            sw.Write("/>");
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFTextInput
    {

        private CT_FFTextType typeField;

        private CT_String defaultField;

        private CT_DecimalNumber maxLengthField;

        private CT_String formatField;

        public CT_FFTextInput()
        {
            this.formatField = new CT_String();
            this.maxLengthField = new CT_DecimalNumber();
            this.defaultField = new CT_String();
            this.typeField = new CT_FFTextType();
        }

        [XmlElement(Order = 0)]
        public CT_FFTextType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_String @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_DecimalNumber maxLength
        {
            get
            {
                return this.maxLengthField;
            }
            set
            {
                this.maxLengthField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_String format
        {
            get
            {
                return this.formatField;
            }
            set
            {
                this.formatField = value;
            }
        }

        public static CT_FFTextInput Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FFTextInput ctObj = new CT_FFTextInput();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "type")
                {
                    ctObj.typeField = CT_FFTextType.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "default")
                {
                    ctObj.defaultField = CT_String.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "format")
                {
                    ctObj.formatField = CT_String.Parse(childNode, namespaceManager);
                }
                else if (childNode.LocalName == "maxLength")
                {
                    ctObj.maxLengthField = CT_DecimalNumber.Parse(childNode, namespaceManager);
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            if (this.typeField == null)
                this.typeField.Write(sw, "w:type");
            if (this.defaultField != null)
                this.defaultField.Write(sw, "w:default");
            if (this.formatField != null)
                this.formatField.Write(sw, "w:format");
            if (this.maxLengthField != null)
                this.maxLengthField.Write(sw, "w:maxLength");
            
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum FFDataItemsType
    {

    
        calcOnExit,

    
        checkBox,

    
        ddList,

    
        enabled,

    
        entryMacro,

    
        exitMacro,

    
        helpText,

    
        name,

    
        statusText,

    
        textInput,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_FldCharType
    {

    
        begin,

    
        separate,

    
        end,
    }

}
