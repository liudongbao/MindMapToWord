using System;
using System.Windows.Forms;

namespace MindMapToWord
{
    /// <summary>
    /// 批量转换进度对话框
    /// </summary>
    public class BatchProgressForm : Form
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _lblCurrent;
        private readonly Label _lblCount;
        private readonly Button _btnCancel;
        private readonly int _total;

        public event EventHandler? CancelRequested;

        public BatchProgressForm(int totalFiles)
        {
            _total = totalFiles;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            Width = 520;
            Height = 160;
            Font = new System.Drawing.Font("微软雅黑", 9f);

            _lblCount = new Label
            {
                Text = $"共 {totalFiles} 个文件",
                AutoSize = true,
                Top = 18,
                Left = 20
            };

            _lblCurrent = new Label
            {
                Text = "正在处理：",
                AutoSize = true,
                Top = 48,
                Left = 20,
                MaximumSize = new System.Drawing.Size(480, 0)
            };

            _progressBar = new ProgressBar
            {
                Top = 78,
                Left = 20,
                Width = 460,
                Height = 22,
                Minimum = 0,
                Maximum = totalFiles,
                Value = 0,
                Style = ProgressBarStyle.Blocks
            };

            _btnCancel = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 28,
                Top = 18,
                Left = 400,
                FlatStyle = FlatStyle.System
            };
            _btnCancel.Click += (_, _) => CancelRequested?.Invoke(this, EventArgs.Empty);

            Controls.AddRange(new Control[] { _lblCount, _lblCurrent, _progressBar, _btnCancel });
        }

        public void UpdateProgress(int current, string currentFileName)
        {
            if (_progressBar.InvokeRequired)
            {
                _progressBar.Invoke(() => UpdateProgress(current, currentFileName));
                return;
            }

            _progressBar.Value = Math.Min(current, _total);
            _lblCurrent.Text = $"正在处理：{currentFileName}";
            _lblCount.Text = $"进度：{current} / {_total}";
        }
    }
}
