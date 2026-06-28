using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace DeepExcel.AddIn
{
    /// <summary>
    /// 任务面板 UserControl - 宿主在 Excel CustomTaskPane 中
    /// 必须标记 ComVisible + ProgId 以便 Excel 通过 ICTPFactory.CreateCTP 实例化
    /// </summary>
    [ComVisible(true)]
    [Guid("B2C3D4E5-F6A7-404B-9A5F-9B3C1D2E3F4B")]
    [ProgId("DeepExcel.AddIn.TaskPaneControl")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITaskPaneControlHost))]
    public class TaskPaneControl : UserControl, ITaskPaneControlHost
    {
        public WebView2 WebView { get; private set; }

        public TaskPaneControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Name = "TaskPaneControl";
            this.Dock = DockStyle.Fill;
            this.BackColor = System.Drawing.Color.White;

            WebView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(WebView);

            this.ResumeLayout(false);
        }
    }

    /// <summary>
    /// COM 默认接口（CTP 宿主需要一个 IDispatch 接口）
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITaskPaneControlHost
    {
        // 空接口，仅为满足 [ComDefaultInterface] 要求
    }
}
