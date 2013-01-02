using System.Windows.Forms;

namespace NPOI.Tools.POIFSBrowser
{
    public partial class AboutDialog : Form
    {
        private static AboutDialog _instance;
        public static AboutDialog Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new AboutDialog();
                }

                return _instance;
            }
        }

        private AboutDialog()
        {
            InitializeComponent();
        }

        private void webSiteLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(webSiteLink.Text);
        }

        private class EtchedBorderedPanel : Panel
        {
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                ControlPaint.DrawBorder3D(e.Graphics, this.ClientRectangle, Border3DStyle.Etched, Border3DSide.Top);
            }
        }
    }
}