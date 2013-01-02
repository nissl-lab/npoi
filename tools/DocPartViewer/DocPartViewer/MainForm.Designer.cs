namespace DocPartViewer
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuOpenFile = new System.Windows.Forms.ToolStripMenuItem();
            this.menuOpenAnother = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.menuExit = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.splitTree = new System.Windows.Forms.SplitContainer();
            this.treeDocPart1 = new System.Windows.Forms.TreeView();
            this.treeDocPart2 = new System.Windows.Forms.TreeView();
            this.splitEditor = new System.Windows.Forms.SplitContainer();
            this.txtEditor1 = new ICSharpCode.TextEditor.TextEditorControl();
            this.txtEditor2 = new ICSharpCode.TextEditor.TextEditorControl();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitTree)).BeginInit();
            this.splitTree.Panel1.SuspendLayout();
            this.splitTree.Panel2.SuspendLayout();
            this.splitTree.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitEditor)).BeginInit();
            this.splitEditor.Panel1.SuspendLayout();
            this.splitEditor.Panel2.SuspendLayout();
            this.splitEditor.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(773, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuOpenFile,
            this.menuOpenAnother,
            this.toolStripSeparator1,
            this.menuExit});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(39, 21);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // menuOpenFile
            // 
            this.menuOpenFile.Name = "menuOpenFile";
            this.menuOpenFile.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.menuOpenFile.Size = new System.Drawing.Size(201, 22);
            this.menuOpenFile.Text = "&Open";
            this.menuOpenFile.Click += new System.EventHandler(this.menuOpenFile_Click);
            // 
            // menuOpenAnother
            // 
            this.menuOpenAnother.Name = "menuOpenAnother";
            this.menuOpenAnother.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.L)));
            this.menuOpenAnother.Size = new System.Drawing.Size(201, 22);
            this.menuOpenAnother.Text = "Open Another";
            this.menuOpenAnother.Click += new System.EventHandler(this.menuOpenAnother_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(198, 6);
            // 
            // menuExit
            // 
            this.menuExit.Name = "menuExit";
            this.menuExit.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.menuExit.Size = new System.Drawing.Size(201, 22);
            this.menuExit.Text = "&Exit";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 428);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(773, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 25);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(773, 25);
            this.toolStrip1.TabIndex = 2;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton1.Text = "toolStripButton1";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 50);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.splitTree);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitEditor);
            this.splitContainer1.Size = new System.Drawing.Size(773, 378);
            this.splitContainer1.SplitterDistance = 257;
            this.splitContainer1.TabIndex = 3;
            // 
            // splitTree
            // 
            this.splitTree.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitTree.Location = new System.Drawing.Point(0, 0);
            this.splitTree.Name = "splitTree";
            this.splitTree.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitTree.Panel1
            // 
            this.splitTree.Panel1.Controls.Add(this.treeDocPart1);
            // 
            // splitTree.Panel2
            // 
            this.splitTree.Panel2.Controls.Add(this.treeDocPart2);
            this.splitTree.Size = new System.Drawing.Size(257, 378);
            this.splitTree.SplitterDistance = 189;
            this.splitTree.TabIndex = 1;
            // 
            // treeDocPart1
            // 
            this.treeDocPart1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeDocPart1.Location = new System.Drawing.Point(0, 0);
            this.treeDocPart1.Name = "treeDocPart1";
            this.treeDocPart1.Size = new System.Drawing.Size(257, 189);
            this.treeDocPart1.TabIndex = 1;
            this.treeDocPart1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeDocPart_AfterSelect);
            // 
            // treeDocPart2
            // 
            this.treeDocPart2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeDocPart2.Location = new System.Drawing.Point(0, 0);
            this.treeDocPart2.Name = "treeDocPart2";
            this.treeDocPart2.Size = new System.Drawing.Size(257, 185);
            this.treeDocPart2.TabIndex = 0;
            this.treeDocPart2.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeDocPart_AfterSelect);
            // 
            // splitEditor
            // 
            this.splitEditor.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitEditor.Location = new System.Drawing.Point(0, 0);
            this.splitEditor.Name = "splitEditor";
            this.splitEditor.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitEditor.Panel1
            // 
            this.splitEditor.Panel1.Controls.Add(this.txtEditor1);
            // 
            // splitEditor.Panel2
            // 
            this.splitEditor.Panel2.Controls.Add(this.txtEditor2);
            this.splitEditor.Size = new System.Drawing.Size(512, 378);
            this.splitEditor.SplitterDistance = 189;
            this.splitEditor.TabIndex = 1;
            // 
            // txtEditor1
            // 
            this.txtEditor1.AutoScroll = true;
            this.txtEditor1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtEditor1.IsReadOnly = false;
            this.txtEditor1.Location = new System.Drawing.Point(0, 0);
            this.txtEditor1.Name = "txtEditor1";
            this.txtEditor1.Size = new System.Drawing.Size(512, 189);
            this.txtEditor1.TabIndex = 1;
            // 
            // txtEditor2
            // 
            this.txtEditor2.AutoScroll = true;
            this.txtEditor2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtEditor2.IsReadOnly = false;
            this.txtEditor2.Location = new System.Drawing.Point(0, 0);
            this.txtEditor2.Name = "txtEditor2";
            this.txtEditor2.Size = new System.Drawing.Size(512, 185);
            this.txtEditor2.TabIndex = 2;
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Word|*.docx|Excel|*.xlsx|PowerPoint|*.pptx";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(773, 450);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "DocPartViewer";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitTree.Panel1.ResumeLayout(false);
            this.splitTree.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitTree)).EndInit();
            this.splitTree.ResumeLayout(false);
            this.splitEditor.Panel1.ResumeLayout(false);
            this.splitEditor.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitEditor)).EndInit();
            this.splitEditor.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem menuOpenFile;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem menuExit;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SplitContainer splitEditor;
        private ICSharpCode.TextEditor.TextEditorControl txtEditor1;
        private System.Windows.Forms.SplitContainer splitTree;
        private System.Windows.Forms.TreeView treeDocPart1;
        private System.Windows.Forms.TreeView treeDocPart2;
        private ICSharpCode.TextEditor.TextEditorControl txtEditor2;
        private System.Windows.Forms.ToolStripMenuItem menuOpenAnother;

    }
}

