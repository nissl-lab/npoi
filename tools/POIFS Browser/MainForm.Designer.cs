namespace NPOI.Tools.POIFSBrowser
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.fileMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.fileOpenMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileMenSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.fileExitMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.openToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.refreshToolScripButton = new System.Windows.Forms.ToolStripButton();
            this.mainSplitContainer = new System.Windows.Forms.SplitContainer();
            this.documentsGroupBox = new System.Windows.Forms.GroupBox();
            this.documentTreeView = new System.Windows.Forms.TreeView();
            this.documentTreeContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.saveAsMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.iconsImageList = new System.Windows.Forms.ImageList(this.components);
            this.streamGroupBox = new System.Windows.Forms.GroupBox();
            this.streamTabControl = new System.Windows.Forms.TabControl();
            this.tabPageBinary = new System.Windows.Forms.TabPage();
            this.viewableTextBox = new System.Windows.Forms.TextBox();
            this.tabPageProperties = new System.Windows.Forms.TabPage();
            this.propertiesListView = new System.Windows.Forms.ListView();
            this.columnId = new System.Windows.Forms.ColumnHeader();
            this.columnName = new System.Windows.Forms.ColumnHeader();
            this.columnType = new System.Windows.Forms.ColumnHeader();
            this.columnValue = new System.Windows.Forms.ColumnHeader();
            this.menuStrip.SuspendLayout();
            this.toolStrip.SuspendLayout();
            this.mainSplitContainer.Panel1.SuspendLayout();
            this.mainSplitContainer.Panel2.SuspendLayout();
            this.mainSplitContainer.SuspendLayout();
            this.documentsGroupBox.SuspendLayout();
            this.documentTreeContextMenu.SuspendLayout();
            this.streamGroupBox.SuspendLayout();
            this.streamTabControl.SuspendLayout();
            this.tabPageBinary.SuspendLayout();
            this.tabPageProperties.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip
            // 
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileMenu,
            this.helpToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(784, 25);
            this.menuStrip.TabIndex = 0;
            this.menuStrip.Text = "menuStrip1";
            // 
            // fileMenu
            // 
            this.fileMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileOpenMenuItem,
            this.fileMenSeparator1,
            this.fileExitMenuItem});
            this.fileMenu.Name = "fileMenu";
            this.fileMenu.Size = new System.Drawing.Size(39, 21);
            this.fileMenu.Text = "&File";
            // 
            // fileOpenMenuItem
            // 
            this.fileOpenMenuItem.Image = global::NPOI.Tools.POIFSBrowser.Properties.Resources.Open;
            this.fileOpenMenuItem.Name = "fileOpenMenuItem";
            this.fileOpenMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.fileOpenMenuItem.Size = new System.Drawing.Size(164, 22);
            this.fileOpenMenuItem.Text = "&Open...";
            this.fileOpenMenuItem.Click += new System.EventHandler(this.Open_Click);
            // 
            // fileMenSeparator1
            // 
            this.fileMenSeparator1.Name = "fileMenSeparator1";
            this.fileMenSeparator1.Size = new System.Drawing.Size(161, 6);
            // 
            // fileExitMenuItem
            // 
            this.fileExitMenuItem.Name = "fileExitMenuItem";
            this.fileExitMenuItem.Size = new System.Drawing.Size(164, 22);
            this.fileExitMenuItem.Text = "E&xit";
            this.fileExitMenuItem.Click += new System.EventHandler(this.fileExitMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(47, 21);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // aboutMenuItem
            // 
            this.aboutMenuItem.Name = "aboutMenuItem";
            this.aboutMenuItem.Size = new System.Drawing.Size(120, 22);
            this.aboutMenuItem.Text = "&About...";
            this.aboutMenuItem.Click += new System.EventHandler(this.aboutMenuItem_Click);
            // 
            // statusStrip
            // 
            this.statusStrip.Location = new System.Drawing.Point(0, 542);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(784, 22);
            this.statusStrip.TabIndex = 1;
            this.statusStrip.Text = "statusStrip1";
            // 
            // toolStrip
            // 
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripButton,
            this.refreshToolScripButton});
            this.toolStrip.Location = new System.Drawing.Point(0, 25);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(784, 25);
            this.toolStrip.TabIndex = 2;
            this.toolStrip.Text = "toolStrip1";
            // 
            // openToolStripButton
            // 
            this.openToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.openToolStripButton.Image = global::NPOI.Tools.POIFSBrowser.Properties.Resources.Open;
            this.openToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openToolStripButton.Name = "openToolStripButton";
            this.openToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.openToolStripButton.Text = "Open File";
            this.openToolStripButton.ToolTipText = "Open Document";
            this.openToolStripButton.Click += new System.EventHandler(this.Open_Click);
            // 
            // refreshToolScripButton
            // 
            this.refreshToolScripButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.refreshToolScripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.refreshToolScripButton.Image = global::NPOI.Tools.POIFSBrowser.Properties.Resources.arrow_refresh;
            this.refreshToolScripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.refreshToolScripButton.Name = "refreshToolScripButton";
            this.refreshToolScripButton.Size = new System.Drawing.Size(23, 22);
            this.refreshToolScripButton.Text = "Refresh File";
            this.refreshToolScripButton.Click += new System.EventHandler(this.refreshToolScripButton_Click);
            // 
            // mainSplitContainer
            // 
            this.mainSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainSplitContainer.Location = new System.Drawing.Point(0, 50);
            this.mainSplitContainer.Name = "mainSplitContainer";
            // 
            // mainSplitContainer.Panel1
            // 
            this.mainSplitContainer.Panel1.Controls.Add(this.documentsGroupBox);
            this.mainSplitContainer.Panel1.Padding = new System.Windows.Forms.Padding(3, 0, 0, 0);
            // 
            // mainSplitContainer.Panel2
            // 
            this.mainSplitContainer.Panel2.Controls.Add(this.streamGroupBox);
            this.mainSplitContainer.Panel2.Padding = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.mainSplitContainer.Size = new System.Drawing.Size(784, 492);
            this.mainSplitContainer.SplitterDistance = 261;
            this.mainSplitContainer.TabIndex = 3;
            // 
            // documentsGroupBox
            // 
            this.documentsGroupBox.Controls.Add(this.documentTreeView);
            this.documentsGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.documentsGroupBox.Location = new System.Drawing.Point(3, 0);
            this.documentsGroupBox.Name = "documentsGroupBox";
            this.documentsGroupBox.Padding = new System.Windows.Forms.Padding(3, 3, 3, 6);
            this.documentsGroupBox.Size = new System.Drawing.Size(258, 492);
            this.documentsGroupBox.TabIndex = 1;
            this.documentsGroupBox.TabStop = false;
            this.documentsGroupBox.Text = "Documents";
            // 
            // documentTreeView
            // 
            this.documentTreeView.ContextMenuStrip = this.documentTreeContextMenu;
            this.documentTreeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.documentTreeView.ImageIndex = 0;
            this.documentTreeView.ImageList = this.iconsImageList;
            this.documentTreeView.Location = new System.Drawing.Point(3, 17);
            this.documentTreeView.Name = "documentTreeView";
            this.documentTreeView.SelectedImageIndex = 0;
            this.documentTreeView.Size = new System.Drawing.Size(252, 469);
            this.documentTreeView.TabIndex = 0;
            this.documentTreeView.AfterCollapse += new System.Windows.Forms.TreeViewEventHandler(this.documentTreeView_AfterCollapse);
            this.documentTreeView.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.documentTreeView_AfterSelect);
            this.documentTreeView.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.documentTreeView_AfterExpand);
            // 
            // documentTreeContextMenu
            // 
            this.documentTreeContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveAsMenu});
            this.documentTreeContextMenu.Name = "documentTreeContextMenu";
            this.documentTreeContextMenu.Size = new System.Drawing.Size(130, 26);
            // 
            // saveAsMenu
            // 
            this.saveAsMenu.Image = global::NPOI.Tools.POIFSBrowser.Properties.Resources.SaveAs;
            this.saveAsMenu.Name = "saveAsMenu";
            this.saveAsMenu.Size = new System.Drawing.Size(129, 22);
            this.saveAsMenu.Text = "Save as...";
            this.saveAsMenu.Click += new System.EventHandler(this.saveAsMenu_Click);
            // 
            // iconsImageList
            // 
            this.iconsImageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("iconsImageList.ImageStream")));
            this.iconsImageList.TransparentColor = System.Drawing.Color.Transparent;
            this.iconsImageList.Images.SetKeyName(0, "File");
            this.iconsImageList.Images.SetKeyName(1, "Folder");
            this.iconsImageList.Images.SetKeyName(2, "FolderOpen");
            this.iconsImageList.Images.SetKeyName(3, "Property");
            this.iconsImageList.Images.SetKeyName(4, "Binary");
            this.iconsImageList.Images.SetKeyName(5, "SummaryStream");
            // 
            // streamGroupBox
            // 
            this.streamGroupBox.Controls.Add(this.streamTabControl);
            this.streamGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.streamGroupBox.Location = new System.Drawing.Point(0, 0);
            this.streamGroupBox.Name = "streamGroupBox";
            this.streamGroupBox.Padding = new System.Windows.Forms.Padding(3, 3, 3, 6);
            this.streamGroupBox.Size = new System.Drawing.Size(516, 492);
            this.streamGroupBox.TabIndex = 1;
            this.streamGroupBox.TabStop = false;
            this.streamGroupBox.Text = "Stream";
            // 
            // streamTabControl
            // 
            this.streamTabControl.Controls.Add(this.tabPageBinary);
            this.streamTabControl.Controls.Add(this.tabPageProperties);
            this.streamTabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.streamTabControl.ImageList = this.iconsImageList;
            this.streamTabControl.Location = new System.Drawing.Point(3, 17);
            this.streamTabControl.Name = "streamTabControl";
            this.streamTabControl.SelectedIndex = 0;
            this.streamTabControl.Size = new System.Drawing.Size(510, 469);
            this.streamTabControl.TabIndex = 2;
            // 
            // tabPageBinary
            // 
            this.tabPageBinary.Controls.Add(this.viewableTextBox);
            this.tabPageBinary.ImageKey = "Binary";
            this.tabPageBinary.Location = new System.Drawing.Point(4, 23);
            this.tabPageBinary.Name = "tabPageBinary";
            this.tabPageBinary.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBinary.Size = new System.Drawing.Size(502, 442);
            this.tabPageBinary.TabIndex = 0;
            this.tabPageBinary.Text = "Binary";
            this.tabPageBinary.UseVisualStyleBackColor = true;
            // 
            // viewableTextBox
            // 
            this.viewableTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.viewableTextBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.viewableTextBox.Location = new System.Drawing.Point(3, 3);
            this.viewableTextBox.Multiline = true;
            this.viewableTextBox.Name = "viewableTextBox";
            this.viewableTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.viewableTextBox.Size = new System.Drawing.Size(496, 436);
            this.viewableTextBox.TabIndex = 1;
            this.viewableTextBox.WordWrap = false;
            // 
            // tabPageProperties
            // 
            this.tabPageProperties.Controls.Add(this.propertiesListView);
            this.tabPageProperties.ImageKey = "Property";
            this.tabPageProperties.Location = new System.Drawing.Point(4, 23);
            this.tabPageProperties.Name = "tabPageProperties";
            this.tabPageProperties.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageProperties.Size = new System.Drawing.Size(502, 442);
            this.tabPageProperties.TabIndex = 1;
            this.tabPageProperties.Text = "Properties";
            this.tabPageProperties.UseVisualStyleBackColor = true;
            // 
            // propertiesListView
            // 
            this.propertiesListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnId,
            this.columnName,
            this.columnType,
            this.columnValue});
            this.propertiesListView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.propertiesListView.FullRowSelect = true;
            this.propertiesListView.GridLines = true;
            this.propertiesListView.Location = new System.Drawing.Point(3, 3);
            this.propertiesListView.Name = "propertiesListView";
            this.propertiesListView.ShowItemToolTips = true;
            this.propertiesListView.Size = new System.Drawing.Size(496, 436);
            this.propertiesListView.SmallImageList = this.iconsImageList;
            this.propertiesListView.TabIndex = 0;
            this.propertiesListView.UseCompatibleStateImageBehavior = false;
            this.propertiesListView.View = System.Windows.Forms.View.Details;
            // 
            // columnId
            // 
            this.columnId.Text = "ID";
            this.columnId.Width = 49;
            // 
            // columnName
            // 
            this.columnName.Text = "Name";
            this.columnName.Width = 150;
            // 
            // columnType
            // 
            this.columnType.Text = "Type";
            this.columnType.Width = 59;
            // 
            // columnValue
            // 
            this.columnValue.Text = "Value";
            this.columnValue.Width = 225;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 564);
            this.Controls.Add(this.mainSplitContainer);
            this.Controls.Add(this.toolStrip);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.menuStrip);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.MainMenuStrip = this.menuStrip;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "POIFS Browser";
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.mainSplitContainer.Panel1.ResumeLayout(false);
            this.mainSplitContainer.Panel2.ResumeLayout(false);
            this.mainSplitContainer.ResumeLayout(false);
            this.documentsGroupBox.ResumeLayout(false);
            this.documentTreeContextMenu.ResumeLayout(false);
            this.streamGroupBox.ResumeLayout(false);
            this.streamTabControl.ResumeLayout(false);
            this.tabPageBinary.ResumeLayout(false);
            this.tabPageBinary.PerformLayout();
            this.tabPageProperties.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem fileMenu;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.SplitContainer mainSplitContainer;
        private System.Windows.Forms.TreeView documentTreeView;
        private System.Windows.Forms.ImageList iconsImageList;
        private System.Windows.Forms.ToolStripMenuItem fileOpenMenuItem;
        private System.Windows.Forms.ToolStripButton openToolStripButton;
        private System.Windows.Forms.ListView propertiesListView;
        private System.Windows.Forms.ColumnHeader columnId;
        private System.Windows.Forms.ColumnHeader columnName;
        private System.Windows.Forms.ColumnHeader columnType;
        private System.Windows.Forms.ColumnHeader columnValue;
        private System.Windows.Forms.GroupBox streamGroupBox;
        private System.Windows.Forms.GroupBox documentsGroupBox;
        private System.Windows.Forms.TextBox viewableTextBox;
        private System.Windows.Forms.ContextMenuStrip documentTreeContextMenu;
        private System.Windows.Forms.ToolStripMenuItem saveAsMenu;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutMenuItem;
        private System.Windows.Forms.TabControl streamTabControl;
        private System.Windows.Forms.TabPage tabPageBinary;
        private System.Windows.Forms.TabPage tabPageProperties;
        private System.Windows.Forms.ToolStripSeparator fileMenSeparator1;
        private System.Windows.Forms.ToolStripMenuItem fileExitMenuItem;
        private System.Windows.Forms.ToolStripButton refreshToolScripButton;
    }
}

