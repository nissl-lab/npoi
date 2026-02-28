namespace NPOI.StyleViewer
{
    partial class Default
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
            this.tvWordStyle = new System.Windows.Forms.TreeView();
            this.rtxtStyleXml = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // tvWordStyle
            // 
            this.tvWordStyle.Dock = System.Windows.Forms.DockStyle.Left;
            this.tvWordStyle.Location = new System.Drawing.Point(0, 0);
            this.tvWordStyle.Name = "tvWordStyle";
            this.tvWordStyle.Size = new System.Drawing.Size(326, 606);
            this.tvWordStyle.TabIndex = 0;
            this.tvWordStyle.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvWordStyle_AfterSelect);
            // 
            // rtxtStyleXml
            // 
            this.rtxtStyleXml.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rtxtStyleXml.Location = new System.Drawing.Point(326, 0);
            this.rtxtStyleXml.Name = "rtxtStyleXml";
            this.rtxtStyleXml.Size = new System.Drawing.Size(782, 606);
            this.rtxtStyleXml.TabIndex = 1;
            this.rtxtStyleXml.Text = "";
            // 
            // Default
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1108, 606);
            this.Controls.Add(this.rtxtStyleXml);
            this.Controls.Add(this.tvWordStyle);
            this.Name = "Default";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NPOI 2.0  StyleViewer";
            this.Load += new System.EventHandler(this.Default_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView tvWordStyle;
        private System.Windows.Forms.RichTextBox rtxtStyleXml;
    }
}