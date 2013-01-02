using System;
using System.Collections;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

using NPOI.HSSF.Record;
using NPOI.SS.Formula.PTG;

namespace NPOI.Tools.POIFSBrowser
{
    public abstract class AbstractRecordTreeNode:TreeNode
    {
        public abstract bool HasBinary { get; }

        public object Record { get; protected set; }

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
                        //parse Ptg array
                        if (propertyValue is Ptg[])
                        {
                            propertyValueText = "";
                            Ptg[] ptgs=(Ptg[])propertyValue;
                            for(int i=0;i<ptgs.Length;i++)
                            {
                                propertyValueText += ptgs[i].ToString();
                            }
                        }
                        else
                        {
                            propertyValueText = propertyValue.ToString();
                        }
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


        public abstract byte[] GetBytes();
    }
}
