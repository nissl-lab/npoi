using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocPartViewer
{
    public enum ZipEntryType
    {
        File, Directory
    }
    public class ZipEntryData : IComparer<ZipEntryData>, IComparable<ZipEntryData>
    {
        public ZipEntryType Type
        {
            get;
            set;
        }
        public string Name { get; set; }
        public string Content { get; set; }
        private List<ZipEntryData> childData = new List<ZipEntryData>();
        public List<ZipEntryData> ChildData
        {
            get { return childData; }
        }

        public int Compare(ZipEntryData x, ZipEntryData y)
        {
            if (x == null && y != null)
                return -1;
            if (x != null && y == null)
                return 1;
            if (x == null && y == null)
                return 0;
            if (x.Type == ZipEntryType.Directory && y.Type == ZipEntryType.File)
                return -1;
            if (x.Type == ZipEntryType.File && y.Type == ZipEntryType.Directory)
                return 1;

            return string.Compare(x.Name, y.Name);
        }

        public int CompareTo(ZipEntryData other)
        {
            return Compare(this, other);
        }
    }
}
