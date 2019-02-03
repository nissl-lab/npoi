using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;
using System.IO;
using NPOI.OpenXml4Net.Util;
using System.Xml;

namespace NPOI.OpenXmlFormats {
    
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_Connections {
        
        private CT_Connection[] connectionField;
        
    
        [XmlElement("connection")]
        public CT_Connection[] connection {
            get {
                return this.connectionField;
            }
            set {
                this.connectionField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_Connection {
        
        private CT_DbPr dbPrField;
        
        private CT_OlapPr olapPrField;
        
        private CT_WebPr webPrField;
        
        private CT_TextPr textPrField;
        
        private CT_Parameters parametersField;
        
        private CT_ExtensionList extLstField;
        
        private uint idField;
        
        private string sourceFileField;
        
        private string odcFileField;
        
        private bool keepAliveField;
        
        private uint intervalField;
        
        private string nameField;
        
        private string descriptionField;
        
        private uint typeField;
        
        private bool typeFieldSpecified;
        
        private uint reconnectionMethodField;
        
        private byte refreshedVersionField;
        
        private byte minRefreshableVersionField;
        
        private bool savePasswordField;
        
        private bool newField;
        
        private bool deletedField;
        
        private bool onlyUseConnectionFileField;
        
        private bool backgroundField;
        
        private bool refreshOnLoadField;
        
        private bool saveDataField;
        
        private ST_CredMethod credentialsField;
        
        private string singleSignOnIdField;
        
        public CT_Connection() {
            this.keepAliveField = false;
            this.intervalField = ((uint)(0));
            this.reconnectionMethodField = ((uint)(1));
            this.minRefreshableVersionField = ((byte)(0));
            this.savePasswordField = false;
            this.newField = false;
            this.deletedField = false;
            this.onlyUseConnectionFileField = false;
            this.backgroundField = false;
            this.refreshOnLoadField = false;
            this.saveDataField = false;
            this.credentialsField = ST_CredMethod.integrated;
        }
        
    
        public CT_DbPr dbPr {
            get {
                return this.dbPrField;
            }
            set {
                this.dbPrField = value;
            }
        }
        
    
        public CT_OlapPr olapPr {
            get {
                return this.olapPrField;
            }
            set {
                this.olapPrField = value;
            }
        }
        
    
        public CT_WebPr webPr {
            get {
                return this.webPrField;
            }
            set {
                this.webPrField = value;
            }
        }
        
    
        public CT_TextPr textPr {
            get {
                return this.textPrField;
            }
            set {
                this.textPrField = value;
            }
        }
        
    
        public CT_Parameters parameters {
            get {
                return this.parametersField;
            }
            set {
                this.parametersField = value;
            }
        }
        
    
        public CT_ExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
        [XmlAttribute]
        public string sourceFile {
            get {
                return this.sourceFileField;
            }
            set {
                this.sourceFileField = value;
            }
        }
        
    
        [XmlAttribute]
        public string odcFile {
            get {
                return this.odcFileField;
            }
            set {
                this.odcFileField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool keepAlive {
            get {
                return this.keepAliveField;
            }
            set {
                this.keepAliveField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "0")]
        public uint interval {
            get {
                return this.intervalField;
            }
            set {
                this.intervalField = value;
            }
        }
        
    
        [XmlAttribute]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
    
        [XmlAttribute]
        public string description {
            get {
                return this.descriptionField;
            }
            set {
                this.descriptionField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "1")]
        public uint reconnectionMethod {
            get {
                return this.reconnectionMethodField;
            }
            set {
                this.reconnectionMethodField = value;
            }
        }
        
    
        [XmlAttribute]
        public byte refreshedVersion {
            get {
                return this.refreshedVersionField;
            }
            set {
                this.refreshedVersionField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(byte), "0")]
        public byte minRefreshableVersion {
            get {
                return this.minRefreshableVersionField;
            }
            set {
                this.minRefreshableVersionField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool savePassword {
            get {
                return this.savePasswordField;
            }
            set {
                this.savePasswordField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool @new {
            get {
                return this.newField;
            }
            set {
                this.newField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool deleted {
            get {
                return this.deletedField;
            }
            set {
                this.deletedField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool onlyUseConnectionFile {
            get {
                return this.onlyUseConnectionFileField;
            }
            set {
                this.onlyUseConnectionFileField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool background {
            get {
                return this.backgroundField;
            }
            set {
                this.backgroundField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool refreshOnLoad {
            get {
                return this.refreshOnLoadField;
            }
            set {
                this.refreshOnLoadField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool saveData {
            get {
                return this.saveDataField;
            }
            set {
                this.saveDataField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_CredMethod.integrated)]
        public ST_CredMethod credentials {
            get {
                return this.credentialsField;
            }
            set {
                this.credentialsField = value;
            }
        }
        
    
        [XmlAttribute]
        public string singleSignOnId {
            get {
                return this.singleSignOnIdField;
            }
            set {
                this.singleSignOnIdField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DbPr {
        
        private string connectionField;
        
        private string commandField;
        
        private string serverCommandField;
        
        private uint commandTypeField;
        
        public CT_DbPr() {
            this.commandTypeField = ((uint)(2));
        }
        
    
        [XmlAttribute]
        public string connection {
            get {
                return this.connectionField;
            }
            set {
                this.connectionField = value;
            }
        }
        
    
        [XmlAttribute]
        public string command {
            get {
                return this.commandField;
            }
            set {
                this.commandField = value;
            }
        }
        
    
        [XmlAttribute]
        public string serverCommand {
            get {
                return this.serverCommandField;
            }
            set {
                this.serverCommandField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "2")]
        public uint commandType {
            get {
                return this.commandTypeField;
            }
            set {
                this.commandTypeField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Extension {
        
        private System.Xml.XmlElement anyField;
        
        private string uriField;
        
    
        [XmlAnyElement()]
        public System.Xml.XmlElement Any {
            get {
                return this.anyField;
            }
            set {
                this.anyField = value;
            }
        }
        
    
        [XmlAttribute(DataType="token")]
        public string uri {
            get {
                return this.uriField;
            }
            set {
                this.uriField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_ExtensionList {
        
        private CT_Extension[] extField;
        
    
        [XmlElement("ext")]
        public CT_Extension[] ext {
            get {
                return this.extField;
            }
            set {
                this.extField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Index {
        
        private uint vField;

        public static CT_Index Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Index ctObj = new CT_Index();
            ctObj.v = XmlHelper.ReadUInt(node.Attributes["v"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "v", this.v);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        [XmlAttribute]
        public uint v {
            get {
                return this.vField;
            }
            set {
                this.vField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_XStringElement {
        
        private string vField;
        
    
        [XmlAttribute]
        public string v {
            get {
                return this.vField;
            }
            set {
                this.vField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_OlapPr {
        
        private bool localField;
        
        private string localConnectionField;
        
        private bool localRefreshField;
        
        private bool sendLocaleField;
        
        private uint rowDrillCountField;
        
        private bool rowDrillCountFieldSpecified;
        
        private bool serverFillField;
        
        private bool serverNumberFormatField;
        
        private bool serverFontField;
        
        private bool serverFontColorField;
        
        public CT_OlapPr() {
            this.localField = false;
            this.localRefreshField = true;
            this.sendLocaleField = false;
            this.serverFillField = true;
            this.serverNumberFormatField = true;
            this.serverFontField = true;
            this.serverFontColorField = true;
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool local {
            get {
                return this.localField;
            }
            set {
                this.localField = value;
            }
        }
        
    
        [XmlAttribute]
        public string localConnection {
            get {
                return this.localConnectionField;
            }
            set {
                this.localConnectionField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool localRefresh {
            get {
                return this.localRefreshField;
            }
            set {
                this.localRefreshField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool sendLocale {
            get {
                return this.sendLocaleField;
            }
            set {
                this.sendLocaleField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint rowDrillCount {
            get {
                return this.rowDrillCountField;
            }
            set {
                this.rowDrillCountField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool rowDrillCountSpecified {
            get {
                return this.rowDrillCountFieldSpecified;
            }
            set {
                this.rowDrillCountFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool serverFill {
            get {
                return this.serverFillField;
            }
            set {
                this.serverFillField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool serverNumberFormat {
            get {
                return this.serverNumberFormatField;
            }
            set {
                this.serverNumberFormatField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool serverFont {
            get {
                return this.serverFontField;
            }
            set {
                this.serverFontField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool serverFontColor {
            get {
                return this.serverFontColorField;
            }
            set {
                this.serverFontColorField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_WebPr {
        
        private CT_Tables tablesField;
        
        private bool xmlField;
        
        private bool sourceDataField;
        
        private bool parsePreField;
        
        private bool consecutiveField;
        
        private bool firstRowField;
        
        private bool xl97Field;
        
        private bool textDatesField;
        
        private bool xl2000Field;
        
        private string urlField;
        
        private string postField;
        
        private bool htmlTablesField;
        
        private ST_HtmlFmt htmlFormatField;
        
        private string editPageField;
        
        public CT_WebPr() {
            this.xmlField = false;
            this.sourceDataField = false;
            this.parsePreField = false;
            this.consecutiveField = false;
            this.firstRowField = false;
            this.xl97Field = false;
            this.textDatesField = false;
            this.xl2000Field = false;
            this.htmlTablesField = false;
            this.htmlFormatField = ST_HtmlFmt.none;
        }
        
    
        public CT_Tables tables {
            get {
                return this.tablesField;
            }
            set {
                this.tablesField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool xml {
            get {
                return this.xmlField;
            }
            set {
                this.xmlField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool sourceData {
            get {
                return this.sourceDataField;
            }
            set {
                this.sourceDataField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool parsePre {
            get {
                return this.parsePreField;
            }
            set {
                this.parsePreField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool consecutive {
            get {
                return this.consecutiveField;
            }
            set {
                this.consecutiveField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool firstRow {
            get {
                return this.firstRowField;
            }
            set {
                this.firstRowField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool xl97 {
            get {
                return this.xl97Field;
            }
            set {
                this.xl97Field = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool textDates {
            get {
                return this.textDatesField;
            }
            set {
                this.textDatesField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool xl2000 {
            get {
                return this.xl2000Field;
            }
            set {
                this.xl2000Field = value;
            }
        }
        
    
        [XmlAttribute]
        public string url {
            get {
                return this.urlField;
            }
            set {
                this.urlField = value;
            }
        }
        
    
        [XmlAttribute]
        public string post {
            get {
                return this.postField;
            }
            set {
                this.postField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool htmlTables {
            get {
                return this.htmlTablesField;
            }
            set {
                this.htmlTablesField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_HtmlFmt.none)]
        public ST_HtmlFmt htmlFormat {
            get {
                return this.htmlFormatField;
            }
            set {
                this.htmlFormatField = value;
            }
        }
        
    
        [XmlAttribute]
        public string editPage {
            get {
                return this.editPageField;
            }
            set {
                this.editPageField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_Tables {
        
        private List<object> itemsField;
        
        private uint countField;
        
        private bool countFieldSpecified;
        
    
        [XmlElement("m", typeof(CT_TableMissing))]
        [XmlElement("s", typeof(CT_XStringElement))]
        [XmlElement("x", typeof(CT_Index))]
        public List<object> Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint count {
            get {
                return this.countField;
            }
            set {
                this.countField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool countSpecified {
            get {
                return this.countFieldSpecified;
            }
            set {
                this.countFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_TableMissing {
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_HtmlFmt {
        
    
        none,
        
    
        rtf,
        
    
        all,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_TextPr {
        
        private CT_TextFields textFieldsField;
        
        private bool promptField;
        
        private ST_FileType fileTypeField;
        
        private uint codePageField;
        
        private uint firstRowField;
        
        private string sourceFileField;
        
        private bool delimitedField;
        
        private string decimalField;
        
        private string thousandsField;
        
        private bool tabField;
        
        private bool spaceField;
        
        private bool commaField;
        
        private bool semicolonField;
        
        private bool consecutiveField;
        
        private ST_Qualifier qualifierField;
        
        private string delimiterField;
        
        public CT_TextPr() {
            this.promptField = true;
            this.fileTypeField = ST_FileType.win;
            this.codePageField = ((uint)(1252));
            this.firstRowField = ((uint)(1));
            this.sourceFileField = "";
            this.delimitedField = true;
            this.decimalField = ".";
            this.thousandsField = ",";
            this.tabField = true;
            this.spaceField = false;
            this.commaField = false;
            this.semicolonField = false;
            this.consecutiveField = false;
            this.qualifierField = ST_Qualifier.doubleQuote;
        }
        
    
        public CT_TextFields textFields {
            get {
                return this.textFieldsField;
            }
            set {
                this.textFieldsField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool prompt {
            get {
                return this.promptField;
            }
            set {
                this.promptField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_FileType.win)]
        public ST_FileType fileType {
            get {
                return this.fileTypeField;
            }
            set {
                this.fileTypeField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "1252")]
        public uint codePage {
            get {
                return this.codePageField;
            }
            set {
                this.codePageField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "1")]
        public uint firstRow {
            get {
                return this.firstRowField;
            }
            set {
                this.firstRowField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute("")]
        public string sourceFile {
            get {
                return this.sourceFileField;
            }
            set {
                this.sourceFileField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool delimited {
            get {
                return this.delimitedField;
            }
            set {
                this.delimitedField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(".")]
        public string @decimal {
            get {
                return this.decimalField;
            }
            set {
                this.decimalField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(",")]
        public string thousands {
            get {
                return this.thousandsField;
            }
            set {
                this.thousandsField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool tab {
            get {
                return this.tabField;
            }
            set {
                this.tabField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool space {
            get {
                return this.spaceField;
            }
            set {
                this.spaceField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool comma {
            get {
                return this.commaField;
            }
            set {
                this.commaField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool semicolon {
            get {
                return this.semicolonField;
            }
            set {
                this.semicolonField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool consecutive {
            get {
                return this.consecutiveField;
            }
            set {
                this.consecutiveField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_Qualifier.doubleQuote)]
        public ST_Qualifier qualifier {
            get {
                return this.qualifierField;
            }
            set {
                this.qualifierField = value;
            }
        }
        
    
        [XmlAttribute]
        public string delimiter {
            get {
                return this.delimiterField;
            }
            set {
                this.delimiterField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_TextFields {
        
        private CT_TextField[] textFieldField;
        
        private uint countField;
        
        public CT_TextFields() {
            this.countField = ((uint)(1));
        }
        
    
        [XmlElement("textField")]
        public CT_TextField[] textField {
            get {
                return this.textFieldField;
            }
            set {
                this.textFieldField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "1")]
        public uint count {
            get {
                return this.countField;
            }
            set {
                this.countField = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_TextField {
        
        private ST_ExternalConnectionType typeField;
        
        private uint positionField;
        
        public CT_TextField() {
            this.typeField = ST_ExternalConnectionType.general;
            this.positionField = ((uint)(0));
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_ExternalConnectionType.general)]
        public ST_ExternalConnectionType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "0")]
        public uint position {
            get {
                return this.positionField;
            }
            set {
                this.positionField = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_ExternalConnectionType {
        
    
        general,
        
    
        text,
        
    
        MDY,
        
    
        DMY,
        
    
        YMD,
        
    
        MYD,
        
    
        DYM,
        
    
        YDM,
        
    
        skip,
        
    
        EMD,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_FileType {
        
    
        mac,
        
    
        win,
        
    
        dos,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_Qualifier {
        
    
        doubleQuote,
        
    
        singleQuote,
        
    
        none,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_Parameters {
        
        private CT_Parameter[] parameterField;
        
        private uint countField;
        
        private bool countFieldSpecified;
        
    
        [XmlElement("parameter")]
        public CT_Parameter[] parameter {
            get {
                return this.parameterField;
            }
            set {
                this.parameterField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint count {
            get {
                return this.countField;
            }
            set {
                this.countField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool countSpecified {
            get {
                return this.countFieldSpecified;
            }
            set {
                this.countFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_Parameter {
        
        private string nameField;
        
        private int sqlTypeField;
        
        private ST_ParameterType parameterTypeField;
        
        private bool refreshOnChangeField;
        
        private string promptField;
        
        private bool booleanField;
        
        private bool booleanFieldSpecified;
        
        private double doubleField;
        
        private bool doubleFieldSpecified;
        
        private int integerField;
        
        private bool integerFieldSpecified;
        
        private string stringField;
        
        private string cellField;
        
        public CT_Parameter() {
            this.sqlTypeField = 0;
            this.parameterTypeField = ST_ParameterType.prompt;
            this.refreshOnChangeField = false;
        }
        
    
        [XmlAttribute]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(0)]
        public int sqlType {
            get {
                return this.sqlTypeField;
            }
            set {
                this.sqlTypeField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_ParameterType.prompt)]
        public ST_ParameterType parameterType {
            get {
                return this.parameterTypeField;
            }
            set {
                this.parameterTypeField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool refreshOnChange {
            get {
                return this.refreshOnChangeField;
            }
            set {
                this.refreshOnChangeField = value;
            }
        }
        
    
        [XmlAttribute]
        public string prompt {
            get {
                return this.promptField;
            }
            set {
                this.promptField = value;
            }
        }
        
    
        [XmlAttribute]
        public bool boolean {
            get {
                return this.booleanField;
            }
            set {
                this.booleanField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool booleanSpecified {
            get {
                return this.booleanFieldSpecified;
            }
            set {
                this.booleanFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public double @double {
            get {
                return this.doubleField;
            }
            set {
                this.doubleField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool doubleSpecified {
            get {
                return this.doubleFieldSpecified;
            }
            set {
                this.doubleFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public int integer {
            get {
                return this.integerField;
            }
            set {
                this.integerField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool integerSpecified {
            get {
                return this.integerFieldSpecified;
            }
            set {
                this.integerFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public string @string {
            get {
                return this.stringField;
            }
            set {
                this.stringField = value;
            }
        }
        
    
        [XmlAttribute]
        public string cell {
            get {
                return this.cellField;
            }
            set {
                this.cellField = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_ParameterType {
        
    
        prompt,
        
    
        value,
        
    
        cell,
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_CredMethod {
        
    
        integrated,
        
    
        none,
        
    
        stored,
        
    
        prompt,
    }
}
