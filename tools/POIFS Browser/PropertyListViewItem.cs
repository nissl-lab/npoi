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

/* ================================================================
 * POIFS Browser 
 * Author: NPOI Team
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * Huseyin Tufekcilerli     2008.11         1.0
 * Tony Qu                  2009.2.18       1.2 alpha
 * 
 * ==============================================================*/

using System;
using System.Collections;
using System.Windows.Forms;
using NPOI.HPSF;
using NPOI.HPSF.Wellknown;

namespace NPOI.Tools.POIFSBrowser
{
    internal class PropertyListViewItem : ListViewItem, IComparable<PropertyListViewItem>
    {
        public Property Property { get; private set; }
        public string TypeName { get; set; }

        private PropertyListViewItem(Property property, PropertyIDMap propertyIDMap)
            : base(new string[] { 
                property.ID.ToString(), 
                propertyIDMap.ContainsKey(property.ID) ? 
                    propertyIDMap[property.ID].ToString() : string.Empty, 
                property.Type.ToString(), 
                property.Value.ToString()
            })
        {
            this.Property = property;

            this.ImageKey = "Property";
        }

        public static PropertyListViewItem[] Create(PropertySet propertySet)
        {
            PropertyIDMap propertyIDMap;

            if (propertySet.IsDocumentSummaryInformation)
            {
                propertyIDMap = PropertyIDMap.DocumentSummaryInformationProperties;
            }
            else if (propertySet.IsSummaryInformation)
            {
                propertyIDMap = PropertyIDMap.SummaryInformationProperties;
            }
            else
            {
                propertyIDMap = new PropertyIDMap(new Hashtable());
            }

            var length = propertySet.Properties.Length;
            var propertyListViewItems = new PropertyListViewItem[length];
            for (int i = 0; i < length; i++)
            {
                var property = propertySet.Properties[i];

                //TODO:: filter empty name items
                PropertyListViewItem tmp= new PropertyListViewItem(property, propertyIDMap);
                
                propertyListViewItems[i] =tmp ;
            }

            System.Array.Sort(propertyListViewItems);

            return propertyListViewItems;
        }

        #region IComparable<PropertyListViewItem> Members

        public int CompareTo(PropertyListViewItem other)
        {
            return this.Property.ID.CompareTo(other.Property.ID);
        }

        #endregion
    }
}