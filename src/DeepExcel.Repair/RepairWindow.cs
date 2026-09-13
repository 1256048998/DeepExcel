using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace DeepExcel.Repair
{
    /// <summary>
    /// What a user sees after "Excel 里没有 DeepExcel 选项卡".
    ///
    /// The whole point is that they do not have to read anything: the window
    /// diagnoses on open, and if something is wrong the repair button is the
    /// only thing that looks clickable.
    /// </summary>
    public sealed class RepairWindow : Form
    {
        private readonly TextBox _output;
        private readonly Button _repair;
        private readonly Button _bundle;
        private readonly Label _status;

        private RepairWindow()
        {
            Text = "DeepExcel 诊断与修复";
            Width = 760;
            Height = 540;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(600, 400);

            _status = new Label
            {
                Dock = DockStyle.Top,
                Height = 46,
                Padding = new Padding(12, 12, 12, 0),
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 10f, FontStyle.Bold),
                Text = "正在检查…"
            };

            _output = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font(FontFamily.GenericMonospace, 9f),
                BackColor = Color.White,
                Margin = new Padding(12)
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 56,
                Padding = new Padding(12, 10, 12, 10)
            };

            _repair = new Button
            {
                Text = "修复",
                Width = 140,
                Height = 34,
                Enabled = false
            };
            _repair.Click += (s, e) => RunAsync(Engine.Repair);

            _bundle = new Button
            {
                Text = "导出诊断包",
                Width = 140,
                Height = 34,
                Enabled = false
            };
            _bundle.Click += (s, e) => ExportBundle();

            buttons.Controls.Add(_repair);
            buttons.Controls.Add(_bundle);

            Controls.Add(_output);
            Controls.Add(buttons);
            Controls.Add(_status);

            Shown += (s, e) => RunAsync(Engine.Diagnose);
        }

        public static void Run()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new RepairWindow());
        }

        private void RunAsync(Func<Action<string>, Report> work)
        {
            SetBusy(true);
            _output.Clear();

            ThreadPool.QueueUserWorkItem(_ =>
            {
                Report report;
                try
                {
                    report = work(Append);
                }
                catch (Exception ex)
                {
                    Append("执行失败：" + ex.Message);
                    BeginInvoke((Action)(() =>
                    {
                        _status.Text = "检查过程出错";
                        _status.ForeColor = Color.Firebrick;
                        SetBusy(false);
                    }));
                    return;
                }

                BeginInvoke((Action)(() =>
                {
                    if (!report.HasBlocking)
                    {
                        _status.Text = "DeepExcel 状态正常。如果 Excel 中仍看不到选项卡，请完全关闭 Excel 后重新打开。";
                        _status.ForeColor = Color.ForestGreen;
                        _repair.Enabled = false;
                    }
                    else if (report.HasUnrepairableBlocking)
                    {
                        _status.Text = "发现无法自动修复的问题，请导出诊断包发给支持人员。";
                        _status.ForeColor = Color.Firebrick;
                        _repair.Enabled = false;
                    }
                    else
                    {
                        _status.Text = "发现可自动修复的问题，点击「修复」继续。";
                        _status.ForeColor = Color.DarkOrange;
                        _repair.Enabled = true;
                    }
                    _bundle.Enabled = true;
                    UseWaitCursor = false;
                }));
            });
        }

        private void ExportBundle()
        {
            SetBusy(true);
            ThreadPool.QueueUserWorkItem(_ =>
            {
                string path = null;
                string error = null;
                try
                {
                    var report = Engine.Diagnose(Append);
                    path = DiagnosticBundle.Write(AddInIdentity.InstallDirectory, report);
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                }

                BeginInvoke((Action)(() =>
                {
                    SetBusy(false);
                    if (error != null)
                    {
                        MessageBox.Show(this, "导出失败：" + error, "DeepExcel",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    Append("");
                    Append("诊断包已生成：" + path);
                    MessageBox.Show(this,
                        "诊断包已保存到桌面：\r\n\r\n" + System.IO.Path.GetFileName(path) +
                        "\r\n\r\n其中不含单元格内容、工作簿路径或 API Key。",
                        "DeepExcel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            });
        }

        private void SetBusy(bool busy)
        {
            UseWaitCursor = busy;
            _repair.Enabled = false;
            _bundle.Enabled = !busy;
        }

        private void Append(string line)
        {
            if (InvokeRequired)
            {
                BeginInvoke((Action<string>)Append, line);
                return;
            }
            _output.AppendText(line + Environment.NewLine);
        }
    }
}
