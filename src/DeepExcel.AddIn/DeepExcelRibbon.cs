using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;

namespace DeepExcel.AddIn
{
    /// <summary>
    /// DeepExcel Ribbon功能区
    /// </summary>
    [ComVisible(true)]
    public class DeepExcelRibbon : OfficeRibbon
    {
        private RibbonTab _deepExcelTab;
        private RibbonGroup _toolsGroup;
        private RibbonButton _openPanelButton;
        private RibbonButton _helpButton;

        protected override void CreateRibbonObjects()
        {
            base.CreateRibbonObjects();

            // 创建Ribbon Tab
            _deepExcelTab = this.Factory.CreateRibbonTab();
            _deepExcelTab.Label = "DeepExcel";
            _deepExcelTab.ControlId = "tabDeepExcel";

            // 工具分组
            _toolsGroup = this.Factory.CreateRibbonGroup();
            _toolsGroup.Label = "Agent";

            // 打开面板按钮
            _openPanelButton = this.Factory.CreateRibbonButton();
            _openPanelButton.Label = "打开面板";
            _openPanelButton.ControlId = "btnOpenPanel";
            _openPanelButton.Image = CreateButtonImage(Color.FromArgb(0, 120, 212));
            _openPanelButton.Click += OpenPanelButton_Click;

            // 帮助按钮
            _helpButton = this.Factory.CreateRibbonButton();
            _helpButton.Label = "帮助";
            _helpButton.ControlId = "btnHelp";
            _helpButton.Click += HelpButton_Click;

            _toolsGroup.Items.Add(_openPanelButton);
            _toolsGroup.Items.Add(_helpButton);
            _deepExcelTab.Controls.Add(_toolsGroup);

            this.Tabs.Add(_deepExcelTab);
        }

        private void OpenPanelButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.ToggleTaskPane();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Open panel error: {ex}");
            }
        }

        private void HelpButton_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "DeepExcel AI Agent\n\n" +
                "在面板中输入你的Excel需求，AI会自动执行：\n" +
                "• 写公式\n" +
                "• 生成并执行VBA\n" +
                "• 数据清洗\n" +
                "• 创建图表/透视表\n\n" +
                "支持模型: Claude / DeepSeek (可在配置中切换)",
                "DeepExcel 帮助",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private Image CreateButtonImage(Color color)
        {
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var brush = new SolidBrush(color))
                {
                    g.FillRectangle(brush, 2, 2, 12, 12);
                }
            }
            return bmp;
        }
    }
}
