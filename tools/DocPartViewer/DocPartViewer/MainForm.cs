using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;
using ICSharpCode.TextEditor.Document;
using System.Xml;
using ICSharpCode.TextEditor;

namespace DocPartViewer
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            txtEditor1.Document.HighlightingStrategy = HighlightingStrategyFactory.CreateHighlightingStrategy("XML");
            txtEditor1.Encoding = System.Text.Encoding.Default;
            txtEditor2.Document.HighlightingStrategy = HighlightingStrategyFactory.CreateHighlightingStrategy("XML");
            txtEditor2.Encoding = System.Text.Encoding.Default;
            treeDocPart1.Tag = txtEditor1;
            treeDocPart2.Tag = txtEditor2;
        }
        
        List<ZipEntryData> roots = new List<ZipEntryData>();
        Dictionary<string, ZipEntryData> dicDatas = new Dictionary<string, ZipEntryData>();
        private void menuOpenFile_Click(object sender, EventArgs e)
        {
            ShowDocPartTree(treeDocPart1);
        }

        private void ShowDocPartTree(TreeView tv)
        {
            string fileDir = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                roots.Clear();
                dicDatas.Clear();
                string rootFile = " ";
                try
                {
                    //读取压缩文件(zip文件)，准备解压缩  
                    ZipInputStream s = new ZipInputStream(File.OpenRead(openFileDialog.FileName));
                    ZipEntry theEntry;
                    string path = fileDir;
                    //解压出来的文件保存的路径  

                    string rootDir = " ";
                    //根目录下的第一个子文件夹的名称 
                    while ((theEntry = s.GetNextEntry()) != null)
                    {
                        rootDir = Path.GetDirectoryName(theEntry.Name);
                        //得到根目录下的第一级子文件夹的名称  
                        if (rootDir.IndexOf("\\") >= 0)
                        {
                            rootDir = rootDir.Substring(0, rootDir.IndexOf("\\") + 1);
                        }
                        string dir = Path.GetDirectoryName(theEntry.Name);
                        //根目录下的第一级子文件夹的下的文件夹的名称  
                        string fileName = Path.GetFileName(theEntry.Name);
                        //根目录下的文件名称  
                        if (dir != " ")
                        //创建根目录下的子文件夹,不限制级别  
                        {
                            //if (!Directory.Exists(fileDir + "\\" + dir))
                            //{
                            //    path = fileDir + "\\" + dir;
                            //    //在指定的路径创建文件夹  
                            //    Directory.CreateDirectory(path);
                            //}
                            AddDirectory(dir);
                        }
                        else if (dir == " " && fileName != "")
                        //根目录下的文件  
                        {
                            path = fileDir;
                            rootFile = fileName;
                        }
                        else if (dir != " " && fileName != "")
                        //根目录下的第一级子文件夹下的文件  
                        {
                            if (dir.IndexOf("\\") > 0)
                            //指定文件保存的路径  
                            {
                                path = fileDir + "\\" + dir;
                            }
                        }

                        if (dir == rootDir)
                        //判断是不是需要保存在根目录下的文件  
                        {
                            path = fileDir + "\\" + rootDir;
                        }

                        //以下为解压缩zip文件的基本步骤  
                        //基本思路就是遍历压缩文件里的所有文件，创建一个相同的文件。  
                        if (fileName != String.Empty)
                        {
                            string text = GetFileData(s);
                            List<ZipEntryData> childData = FindParentData(dir);
                            ZipEntryData zdata = new ZipEntryData()
                            {
                                Type = ZipEntryType.File,
                                Name = fileName,
                                Content = text
                            };
                            childData.Add(zdata);
                            dicDatas.Add(theEntry.Name, zdata);
                        }
                    }
                    s.Close();

                }
                catch (Exception)
                {
                }
                tv.Nodes.Clear();
                BuildDocPartTree(roots, tv.Nodes);
                tv.ExpandAll();
            }
            
        }

        private void BuildDocPartTree(List<ZipEntryData> nodes, TreeNodeCollection treeNodes)
        {
            nodes.Sort();
            //nodes.Reverse();
            foreach (ZipEntryData data in nodes)
            {
                TreeNode node = new TreeNode();
                node.Text = data.Name;
                node.Tag = data;
                treeNodes.Add(node);
                if (data.Type == ZipEntryType.Directory)
                    BuildDocPartTree(data.ChildData, node.Nodes);
            }
        }

        private static string GetFileData(ZipInputStream s)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                int size = 2048;
                byte[] data = new byte[2048];
                while (true)
                {
                    size = s.Read(data, 0, data.Length);
                    if (size > 0)
                    {
                        ms.Write(data, 0, size);
                    }
                    else
                    {
                        break;
                    }
                }
                string text = Encoding.UTF8.GetString(ms.ToArray());
                ms.Close();
                return text;
            }
        }

        private void AddDirectory(string path)
        {
            List<ZipEntryData> childList = roots;
            string[] dirNames = path.Split(new char[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
            string key = string.Empty;
            ZipEntryData data;
            foreach (string dir in dirNames)
            {
                key = string.IsNullOrEmpty(key) ? dir : key + Path.DirectorySeparatorChar + dir;
                if (dicDatas.ContainsKey(key))
                {
                    data = dicDatas[key];
                    childList = data.ChildData;
                    continue;
                }

                data = new ZipEntryData();
                data.Type = ZipEntryType.Directory;
                data.Name = dir;
                
                dicDatas.Add(key, data);
                childList.Add(data);
                childList = data.ChildData;
            }
        }

        private List<ZipEntryData> FindParentData(string dir)
        {
            if (string.IsNullOrEmpty(dir))
                return roots;
            string key = dir.Trim(Path.DirectorySeparatorChar);
            return dicDatas[key].ChildData;
        }
        private XmlDocument xmlDoc = new XmlDocument();
        private void treeDocPart_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag is ZipEntryData)
            {
                ZipEntryData data = (ZipEntryData)e.Node.Tag;
                if (data.Type == ZipEntryType.File)
                {
                    TextEditorControl editor = e.Node.TreeView.Tag as TextEditorControl;
                    UTF8Encoding utf8 = new UTF8Encoding(false);
                    byte[] test = Encoding.UTF8.GetBytes(data.Content);
                    string xml;
                    //remove utf8 string BOM flags 
                    if (test[0] == 0xef && test[1] == 0xbb && test[2] == 0xbf)
                        xml = utf8.GetString(test, 3, test.Length - 3);
                    else
                        xml = data.Content;
                    xmlDoc.LoadXml(xml);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (XmlTextWriter xmlWriter = new XmlTextWriter(ms, utf8))
                        {
                            xmlWriter.Indentation = 4;
                            xmlWriter.Formatting = System.Xml.Formatting.Indented;

                            xmlDoc.WriteContentTo(xmlWriter);
                            xmlWriter.Close();
                        }
                        string result = Encoding.UTF8.GetString(ms.ToArray());
                        editor.Text = result;
                        editor.Refresh();
                    }
                    
                }
            }
        }

        private void menuOpenAnother_Click(object sender, EventArgs e)
        {
            ShowDocPartTree(treeDocPart2);
        }
    }
}
