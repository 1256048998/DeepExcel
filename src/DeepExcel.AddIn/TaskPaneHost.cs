using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace DeepExcel.AddIn
{
    /// <summary>
    /// 任务面板宿主控件，承载WebView2
    /// </summary>
    public partial class TaskPaneHost : UserControl
    {
        public WebView2 WebView { get; private set; }

        public TaskPaneHost()
        {
            InitializeComponent();
            InitializeWebView();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Name = "TaskPaneHost";
            this.Size = new System.Drawing.Size(420, 600);
            this.BackColor = System.Drawing.Color.White;
            this.ResumeLayout(false);
        }

        private void InitializeWebView()
        {
            WebView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            this.Controls.Add(WebView);
        }
    }
}
