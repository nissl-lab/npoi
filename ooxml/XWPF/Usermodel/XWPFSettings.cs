/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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
namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.OpenXml4Net.OPC;
    using System.IO;
    using System.Xml.Serialization;
    using System.Xml;

    public class XWPFSettings : POIXMLDocumentPart
    {

        private CT_Settings ctSettings;

        public XWPFSettings(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

        }

        public XWPFSettings()
            : base()
        {
            ctSettings = new CT_Settings();
        }


        internal override void OnDocumentRead()
        {
            base.OnDocumentRead();
            ReadFrom(GetPackagePart().GetInputStream());
        }

        /**
         * In the zoom tag inside Settings.xml file <br/>
         * it Sets the value of zoom
         * @return percentage as an integer of zoom level
         */
        public long GetZoomPercent()
        {
            CT_Zoom zoom = ctSettings.zoom;
            if (!ctSettings.IsSetZoom())
            {
                zoom = ctSettings.AddNewZoom();
            }
            else
            {
                zoom = ctSettings.zoom;
            }

            return long.Parse(zoom.percent);
        }

        /// <summary>
        /// Set zoom. In the zoom tag inside settings.xml file it sets the value of zoom
        /// </summary>
        /// <param name="zoomPercent"></param>
        /// <example>
        /// sample snippet from Settings.xml 
        /// 
        /// &lt;w:zoom w:percent="50" /&gt;
        /// </example>
        public void SetZoomPercent(long zoomPercent)
        {
            if (!ctSettings.IsSetZoom())
            {
                ctSettings.AddNewZoom();
            }
            CT_Zoom zoom = ctSettings.zoom;
            zoom.percent = zoomPercent.ToString();
        }

        /**
         * Verifies the documentProtection tag inside Settings.xml file <br/>
         * if the protection is enforced (w:enforcement="1") <br/>
         * and if the kind of protection Equals to passed (STDocProtect.Enum editValue) <br/>
         * 
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;readOnly&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         * 
         * @return true if documentProtection is enforced with option ReadOnly
         */
        public bool IsEnforcedWith(ST_DocProtect editValue)
        {
            CT_DocProtect ctDocProtect = ctSettings.documentProtection;

            if (ctDocProtect == null)
            {
                return false;
            }

            return ctDocProtect.enforcement.Equals(ST_OnOff.on) && ctDocProtect.edit.Equals(editValue);
        }

        /**
         * Enforces the protection with the option specified by passed editValue.<br/>
         * <br/>
         * In the documentProtection tag inside Settings.xml file <br/>
         * it Sets the value of enforcement to "1" (w:enforcement="1") <br/>
         * and the value of edit to the passed editValue (w:edit="[passed editValue]")<br/>
         * <br/>
         * sample snippet from Settings.xml
         * <pre>
         *     &lt;w:settings  ... &gt;
         *         &lt;w:documentProtection w:edit=&quot;[passed editValue]&quot; w:enforcement=&quot;1&quot;/&gt;
         * </pre>
         */
        public void SetEnforcementEditValue(ST_DocProtect editValue)
        {
            SafeGetDocumentProtection().enforcement = (ST_OnOff.on);
            SafeGetDocumentProtection().edit = (editValue);
        }

        /**
         * Removes protection enforcement.<br/>
         * In the documentProtection tag inside Settings.xml file <br/>
         * it Sets the value of enforcement to "0" (w:enforcement="0") <br/>
         */
        public void RemoveEnforcement()
        {
            SafeGetDocumentProtection().enforcement = (ST_OnOff.off);
        }

        /**
         * Enforces fields update on document open (in Word).
         * In the settings.xml file <br/>
         * sets the updateSettings value to true (w:updateSettings w:val="true")
         * 
         *  NOTICES:
         *  <ul>
         *  	<li>Causing Word to ask on open: "This document contains fields that may refer to other files. Do you want to update the fields in this document?"
         *           (if "Update automatic links at open" is enabled)</li>
         *  	<li>Flag is removed after saving with changes in Word </li>
         *  </ul> 
         */
        public void SetUpdateFields()
        {
            CT_OnOff onOff = new CT_OnOff();
            onOff.val = true;
            ctSettings.updateFields=(onOff);
        }

        public bool IsUpdateFields()
        {
            return ctSettings.IsSetUpdateFields() && ctSettings.updateFields.val == true;
        }

        /**
         * get or set revision tracking
         */
        public bool IsTrackRevisions
        {
            get
            {
                return ctSettings.IsSetTrackRevisions();
            }
            set
            {
                if (value)
                {
                    if (!ctSettings.IsSetTrackRevisions())
                    {
                        ctSettings.AddNewTrackRevisions();
                    }
                }
                else
                {
                    if (ctSettings.IsSetTrackRevisions())
                    {
                        ctSettings.UnsetTrackRevisions();
                    }
                }
            }
        }

        protected internal override void Commit()
        {
            if (ctSettings == null)
            {
                throw new InvalidOperationException("Unable to write out settings that were never read in!");
            }
            /*XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.SaveSyntheticDocumentElement=(new QName(CTSettings.type.Name.NamespaceURI, "settings"));
            Dictionary<String, String> map = new Dictionary<String, String>();
            map.Put("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w");
            xmlOptions.SaveSuggestedPrefixes=(map);*/
            //XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            //    new XmlQualifiedName("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")});
            PackagePart part = GetPackagePart();
            using (Stream out1 = part.GetOutputStream())
            {
                SettingsDocument sd = new SettingsDocument(ctSettings);
                sd.Save(out1);
            }
        }

        private CT_DocProtect SafeGetDocumentProtection()
        {
            CT_DocProtect documentProtection = ctSettings.documentProtection;
            if (documentProtection == null)
            {
                documentProtection = new CT_DocProtect();
                ctSettings.documentProtection = (documentProtection);
            }
            return ctSettings.documentProtection;
        }

        private void ReadFrom(Stream inputStream)
        {
            try
            {
                XmlDocument xmldoc = ConvertStreamToXml(inputStream);
                ctSettings = SettingsDocument.Parse(xmldoc,NamespaceManager).Settings;
            }
            catch (Exception e)
            {
                throw new Exception("SettingsDocument parse failed", e);
            }
        }

    }

}