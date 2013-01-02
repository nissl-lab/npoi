using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Reflection;

namespace NPOI.Tools.POIFSBrowser
{
    internal abstract class AbstractHWPFTreeNode : AbstractTreeNode
    {
        public object Record { get; set; }
        public abstract byte[] GetBytes();
        public virtual ListViewItem[] GetPropertyList()
        {
            ArrayList tmplist = new ArrayList();

            PropertyInfo[] properties = Record.GetType().GetProperties();
            int n = 1;
            foreach (PropertyInfo property in properties)
            {
                string propertyValueText = "(null)";
                string propertyValueType = "";
                try
                {
                    object propertyValue = property.GetValue(this.Record, null);


                    if (propertyValue != null)
                    {
                        propertyValueText = propertyValue.ToString();
                        propertyValueType = propertyValue.GetType().Name;
                    }
                }
                catch (Exception e)
                {
                    propertyValueText = e.Message;
                }
                tmplist.Add(new ListViewItem(
                    new string[] { (n++).ToString(), property.Name, propertyValueType, propertyValueText }));
            }
            return (ListViewItem[])tmplist.ToArray(typeof(ListViewItem));
        }
    }
}
