using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MindMapToWord.Core;
using MindMapToWord.Exporters;

namespace MindMapToWord
{
    /// <summary>
    /// 主窗口 —— 选择导图文件 -> 解析预览 -> 导出为 Word
    /// </summary>
    public class MainForm : Form
    {
        private readonly Button _btnOpen = new() { Text = "选择导图文件…", Width = 160, Height = 36 };
        private readonly Label _lblFile = new() { Text = "（尚未选择文件）", AutoSize = false, Dock = DockStyle.Fill, Padding = new Padding(6, 8, 6, 8) };
        private readonly Button _btnPreview = new() { Text = "解析预览", Width = 110, Height = 36, Enabled = false };
        private readonly Button _btnExport = new() { Text = "导出为 Word", Width = 140, Height = 36, Enabled = false };
        private readonly Button _btnBatchExport = new() { Text = "批量导出…", Width = 130, Height = 36, Enabled = false };
        private readonly Button _btnInspect = new() { Text = "🔍 查看内部结构", Width = 150, Height = 36, Enabled = false };
        private readonly CheckBox _chkNotes = new() { Text = "包含备注", Checked = true };
        private readonly CheckBox _chkLinks = new() { Text = "包含超链接", Checked = true };
        private readonly Label _lblStats = new() { AutoSize = true };
        private readonly TreeView _tree = new() { Dock = DockStyle.Fill, ShowLines = true, ShowPlusMinus = true };
        private readonly StatusStrip _status = new();
        private readonly ToolStripStatusLabel _statusLabel = new() { Text = "就绪" };

        private MindMapDocument? _currentDoc;
        private string _currentPath = string.Empty;
        private readonly List<string> _selectedFiles = new();

        public MainForm()
        {
            Text = "思维导图 → Word 转换器";
            Width = 960;
            Height = 660;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("微软雅黑", 9f);
            MinimumSize = new Size(720, 500);

            BuildLayout();
            BindEvents();
        }

        private void BuildLayout()
        {
            // 顶部：文件选择行（Dock 顺序：先右后左，最后 Fill）
            var top = new Panel { Dock = DockStyle.Top, Height = 60, Padding = new Padding(10, 10, 10, 0) };

            var topMid = new Panel { Dock = DockStyle.Fill };
            topMid.Controls.Add(_lblFile);

            var topLeft = new Panel { Dock = DockStyle.Left, Width = 180 };
            topLeft.Controls.Add(_btnOpen);
            _btnOpen.Dock = DockStyle.Fill;

            var topRight = new Panel { Dock = DockStyle.Right, Width = 580 };
            var topRightInner = new Panel { Dock = DockStyle.Fill };
            topRightInner.Controls.Add(_btnBatchExport);
            topRightInner.Controls.Add(_btnExport);
            topRightInner.Controls.Add(_btnInspect);
            topRightInner.Controls.Add(_btnPreview);
            _btnBatchExport.Dock = DockStyle.Right;
            _btnExport.Dock = DockStyle.Right;
            _btnInspect.Dock = DockStyle.Right;
            _btnPreview.Dock = DockStyle.Left;
            topRight.Controls.Add(topRightInner);

            top.Controls.Add(topMid);    // 先加 Fill
            top.Controls.Add(topLeft);   // 后加 Left
            top.Controls.Add(topRight);  // 最后 Right

            // 选项行
            var options = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(10, 0, 10, 0) };
            var flp = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            flp.Controls.Add(_chkNotes);
            flp.Controls.Add(_chkLinks);
            flp.Controls.Add(_lblStats);
            options.Controls.Add(flp);

            // 内容区域（Tree 预览）
            var content = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            content.Controls.Add(_tree);

            // 状态条
            _status.Items.Add(_statusLabel);

            // 按 Dock 顺序添加到主窗体（先加 Fill，再加 Top，最后 Bottom）
            Controls.Add(options);     // Top(较下)
            Controls.Add(top);         // Top(较上)
            Controls.Add(content);     // Fill
            Controls.Add(_status);     // Bottom

            // 让 Top 控件排在 Fill 之上（越后加越靠前）
            options.BringToFront();
            top.BringToFront();
            _status.BringToFront();

            UpdateStats();
        }

        private void BindEvents()
        {
            _btnOpen.Click += BtnOpen_Click;
            _btnPreview.Click += BtnPreview_Click;
            _btnExport.Click += BtnExport_Click;
            _btnBatchExport.Click += BtnBatchExport_Click;
            _btnInspect.Click += BtnInspect_Click;
        }

        private void BtnOpen_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Title = "选择思维导图文件（可多选）",
                Filter = MindMapParserFactory.SupportedFilter,
                CheckFileExists = true,
                Multiselect = true
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            _selectedFiles.Clear();
            _selectedFiles.AddRange(dlg.FileNames);

            if (_selectedFiles.Count == 1)
            {
                _currentPath = _selectedFiles[0];
                _lblFile.Text = _currentPath;
                _btnPreview.Enabled = true;
                _btnInspect.Enabled = true;
            }
            else
            {
                _currentPath = string.Empty;
                _lblFile.Text = $"已选择 {_selectedFiles.Count} 个文件";
                _btnPreview.Enabled = false;
                _btnInspect.Enabled = false;
            }

            _btnExport.Enabled = false;
            _btnBatchExport.Enabled = _selectedFiles.Count > 0;
            _currentDoc = null;
            _tree.Nodes.Clear();
            _statusLabel.Text = $"已选择 {_selectedFiles.Count} 个文件，可点击「批量导出」直接转换。";
            UpdateStats();
        }

        private void BtnPreview_Click(object? sender, EventArgs e)
        {
            try
            {
                _statusLabel.Text = "正在解析，请稍候…";
                Application.DoEvents();

                var doc = MindMapParserFactory.Parse(_currentPath);
                _currentDoc = doc;

                // ====== 树结构诊断 ======
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"文件：{Path.GetFileName(_currentPath)}");
                sb.AppendLine($"根节点数：{doc.Roots.Count}");
                sb.AppendLine($"总节点数：{doc.TotalNodes}");
                sb.AppendLine($"总图片数：{doc.TotalImages}");
                sb.AppendLine();
                sb.AppendLine("【树结构预览】（每行格式：[层级] 「标题」  children:N  order:N）");
                int orderCounter = 0;
                foreach (var root in doc.Roots)
                {
                    DumpNodeWithOrder(sb, root, 0, ref orderCounter);
                }

                var debugForm = new Form
                {
                    Text = "树结构诊断 —— " + Path.GetFileName(_currentPath),
                    Width = 750,
                    Height = 600,
                    StartPosition = FormStartPosition.CenterParent
                };
                var debugTxt = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    Font = new Font("Consolas", 9f),
                    Text = sb.ToString()
                };
                debugForm.Controls.Add(debugTxt);
                debugForm.ShowDialog(this);
                // ====== 诊断结束 ======

                _tree.Nodes.Clear();
                foreach (var root in doc.Roots)
                {
                    var tn = new TreeNode(string.IsNullOrWhiteSpace(root.Title) ? "（根节点）" : root.Title) { Tag = root };
                    PopulateTree(tn, root);
                    _tree.Nodes.Add(tn);
                }
                _tree.ExpandAll();
                if (_tree.Nodes.Count > 0) _tree.SelectedNode = _tree.Nodes[0];

                _btnExport.Enabled = true;
                _statusLabel.Text = $"解析成功，共 {doc.TotalNodes} 个节点、{doc.TotalImages} 张图片。";
                UpdateStats();
            }
            catch (Exception ex)
            {
                MessageBox.Show("解析失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusLabel.Text = "解析失败：" + ex.Message;
            }
        }

        private static void DumpNode(System.Text.StringBuilder sb, MindMapNode node, int depth)
        {
            var indent = new string(' ', depth * 2);
            sb.AppendLine($"{indent}[L{depth}] 「{node.Title}」  children:{node.Children.Count}  images:{node.Images.Count}");
            foreach (var child in node.Children)
                DumpNode(sb, child, depth + 1);
        }

        private static void DumpNodeWithOrder(System.Text.StringBuilder sb, MindMapNode node, int depth, ref int order)
        {
            var indent = new string(' ', depth * 2);
            var yStr = node.RawY != 0 ? $"Y={node.RawY}" : "Y=—";
            var xStr = node.RawX != 0 ? $"X={node.RawX}" : "X=—";
            sb.AppendLine($"{indent}[L{depth}][#{order++}] {yStr} {xStr} 「{node.Title}」 children:{node.Children.Count}");
            if (!string.IsNullOrEmpty(node.RawAttributes))
                sb.AppendLine($"{indent}    attrs: {node.RawAttributes}");
            if (!string.IsNullOrEmpty(node.RawLevelData))
                sb.AppendLine($"{indent}    levelData: {node.RawLevelData}");
            foreach (var child in node.Children)
                DumpNodeWithOrder(sb, child, depth + 1, ref order);
        }

        private static void PopulateTree(TreeNode tn, MindMapNode node)
        {
            var info = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(node.Notes)) info.Append(" · 有备注");
            if (!string.IsNullOrWhiteSpace(node.Hyperlink)) info.Append(" · 有链接");
            if (node.Images.Count > 0) info.Append($" · {node.Images.Count}图");
            if (info.Length > 0) tn.ToolTipText = info.ToString(3, info.Length - 3);

            foreach (var c in node.Children)
            {
                var child = new TreeNode(c.Title) { Tag = c };
                PopulateTree(child, c);
                tn.Nodes.Add(child);
            }
        }

        private void BtnExport_Click(object? sender, EventArgs e)
        {
            if (_currentDoc == null)
            {
                MessageBox.Show("请先选择文件并完成预览解析。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dlg = new SaveFileDialog
            {
                Title = "保存 Word 文档",
                Filter = "Word 文档 (*.docx)|*.docx",
                FileName = Path.GetFileNameWithoutExtension(_currentPath) + ".docx",
                DefaultExt = ".docx",
                AddExtension = true
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                _statusLabel.Text = "正在生成 Word 文档…";
                Application.DoEvents();

                var exporter = new WordExporter(dlg.FileName)
                {
                    IncludeNotes = _chkNotes.Checked,
                    IncludeHyperlinks = _chkLinks.Checked
                };
                exporter.Export(_currentDoc);

                _statusLabel.Text = "导出完成：" + dlg.FileName;
                var result = MessageBox.Show("导出完成！是否立即打开文件？", "成功",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true }); }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusLabel.Text = "导出失败：" + ex.Message;
            }
        }

        private void BtnBatchExport_Click(object? sender, EventArgs e)
        {
            if (_selectedFiles.Count == 0)
            {
                MessageBox.Show("请先选择文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 选择输出目录
            using var folderDlg = new FolderBrowserDialog
            {
                Description = "请选择批量导出的输出目录"
            };
            if (folderDlg.ShowDialog(this) != DialogResult.OK) return;
            var outputDir = folderDlg.SelectedPath;

            // 验证输出目录可写
            try
            {
                var testFile = Path.Combine(outputDir, "_test_write.tmp");
                File.Create(testFile).Dispose();
                File.Delete(testFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法写入目标目录：{ex.Message}\n\n请选择其他目录。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 创建进度对话框
            using var progressForm = new BatchProgressForm(_selectedFiles.Count);
            progressForm.Text = "批量导出进度";
            progressForm.StartPosition = FormStartPosition.CenterParent;

            var includeNotes = _chkNotes.Checked;
            var includeLinks = _chkLinks.Checked;
            var results = new List<(string file, bool success, string message)>();

            // BackgroundWorker 完成批量转换
            var worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += (_, args) =>
            {
                var files = (List<string>)args.Argument!;
                for (int i = 0; i < files.Count; i++)
                {
                    var file = files[i];
                    var fileName = Path.GetFileNameWithoutExtension(file);
                    var outPath = Path.Combine(outputDir, fileName + ".docx");

                    worker.ReportProgress(i + 1, fileName);

                    try
                    {
                        var doc = MindMapParserFactory.Parse(file);
                        var exporter = new WordExporter(outPath)
                        {
                            IncludeNotes = includeNotes,
                            IncludeHyperlinks = includeLinks
                        };
                        exporter.Export(doc);
                        results.Add((fileName, true, outPath));
                    }
                    catch (Exception ex)
                    {
                        results.Add((fileName, false, ex.Message));
                    }
                }
                args.Result = results;
            };

            worker.ProgressChanged += (_, args) =>
            {
                progressForm.UpdateProgress(args.ProgressPercentage, (string)args.UserState!);
            };

            worker.RunWorkerCompleted += (_, args) =>
            {
                progressForm.Close();
                var allResults = (List<(string, bool, string)>)args.Result!;
                ShowBatchSummary(allResults);
            };

            progressForm.CancelRequested += (_, _) => worker.CancelAsync();

            worker.RunWorkerAsync(_selectedFiles.ToList());
            progressForm.ShowDialog(this);
        }

        private void ShowBatchSummary(List<(string file, bool success, string message)> results)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"批量转换完成，共 {results.Count} 个文件：");
            sb.AppendLine();

            var successCount = results.Count(r => r.success);
            var failCount = results.Count(r => !r.success);
            sb.AppendLine($"✅ 成功：{successCount} 个");
            sb.AppendLine($"❌ 失败：{failCount} 个");
            sb.AppendLine();

            if (failCount > 0)
            {
                sb.AppendLine("【失败文件】：");
                foreach (var (file, success, message) in results)
                {
                    if (!success)
                        sb.AppendLine($"  • {file}：{message}");
                }
                sb.AppendLine();
            }

            if (successCount > 0)
            {
                sb.AppendLine("【成功文件】：");
                foreach (var (file, success, message) in results)
                {
                    if (success)
                        sb.AppendLine($"  ✓ {file}");
                }
            }

            MessageBox.Show(sb.ToString(), "批量导出结果",
                MessageBoxButtons.OK, failCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            _statusLabel.Text = $"批量导出完成：{successCount} 成功，{failCount} 失败。";
        }

        private void UpdateStats()
        {
            if (_currentDoc != null)
                _lblStats.Text = $"｜节点：{_currentDoc.TotalNodes}　图片：{_currentDoc.TotalImages}　格式：{_currentDoc.SourceFormat}";
            else
                _lblStats.Text = string.Empty;
        }

        private void BtnInspect_Click(object? sender, EventArgs e)
        {
            try
            {
                _statusLabel.Text = "正在探查文件内部结构…";
                Application.DoEvents();

                var report = new StringBuilder();
                var ext = Path.GetExtension(_currentPath).ToLowerInvariant();
                report.AppendLine($"文件：{Path.GetFileName(_currentPath)}");
                report.AppendLine($"扩展名：{ext}");
                report.AppendLine($"大小：{(new FileInfo(_currentPath).Length / 1024):N0} KB");
                report.AppendLine();

                // ZIP 类（.emmx / .xmind / .mindmap 等）
                if (ext == ".emmx" || ext == ".xmind" || ext == ".mindmap" ||
                    ext == ".mmap" || ext == ".zmindmap" || IsZipFile(_currentPath))
                {
                    using var zip = System.IO.Compression.ZipFile.OpenRead(_currentPath);
                    report.AppendLine($"【ZIP 包内条目】共 {zip.Entries.Count} 个：");
                    int count = 0;
                    foreach (var entry in zip.Entries.OrderByDescending(e => e.Length))
                    {
                        report.AppendLine($"  {entry.Length,8:###,##0} B   {entry.FullName}");
                        count++;
                        if (count > 200) { report.AppendLine("  ...（其余省略）"); break; }
                    }
                    report.AppendLine();

                    // 预览每个 XML/JSON 文件的内容片段（不限大小）
                    foreach (var entry in zip.Entries.OrderByDescending(e => e.Length))
                    {
                        var name = entry.Name.ToLowerInvariant();
                        if (!(name.EndsWith(".json") || name.EndsWith(".xml") ||
                              name.EndsWith(".html") || name == "content" || name == "document"))
                            continue;

                        try
                        {
                            using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
                            var text = sr.ReadToEnd();
                            if (string.IsNullOrWhiteSpace(text)) continue;

                            report.AppendLine($"═══ {entry.FullName} ({entry.Length} B) ═══");

                            if (name.EndsWith(".xml"))
                            {
                                // 截取前 1500 字符作为 header 预览
                                var headerLen = Math.Min(1500, text.Length);
                                report.AppendLine(text.Substring(0, headerLen));
                                report.AppendLine();

                                // 统计关键字
                                var keywordCounts = new Dictionary<string, int>
                                {
                                    ["<Shape"] = CountWord(text, "<Shape"),
                                    ["<shape"] = CountWord(text, "<shape"),
                                    ["<Topic"] = CountWord(text, "<Topic"),
                                    ["<topic"] = CountWord(text, "<topic"),
                                    ["<Node"] = CountWord(text, "<Node"),
                                    ["<node"] = CountWord(text, "<node"),
                                    ["<Text"] = CountWord(text, "<Text"),
                                    ["<text"] = CountWord(text, "<text"),
                                    ["<ap:"] = CountWord(text, "<ap:"),
                                    ["ap:Shape"] = CountWord(text, "ap:Shape"),
                                    ["CentralTopic"] = CountWord(text, "CentralTopic"),
                                    ["MainTopic"] = CountWord(text, "MainTopic"),
                                    ["SubTopic"] = CountWord(text, "SubTopic"),
                                    ["Type="] = CountWord(text, "Type="),
                                    ["Text>"] = CountWord(text, "Text>"),
                                    ["<img"] = CountWord(text, "<img"),
                                    ["<Image"] = CountWord(text, "<Image"),
                                    ["<image"] = CountWord(text, "<image"),
                                };
                                report.AppendLine("  【元素统计】");
                                foreach (var kvp in keywordCounts)
                                {
                                    if (kvp.Value > 0)
                                        report.AppendLine($"    {kvp.Key}: {kvp.Value}");
                                }
                                report.AppendLine();

                                // 找到第一个类似节点的元素，预览 1500 字符
                                var elementKeywords = new[] { "<Shape", "<shape", "<Topic", "<topic", "<Node", "<node", "<ap:" };
                                int foundIdx = -1;
                                foreach (var kw in elementKeywords)
                                {
                                    var pos = text.IndexOf(kw, StringComparison.OrdinalIgnoreCase);
                                    if (pos >= 0) { foundIdx = pos; break; }
                                }
                                if (foundIdx > 0)
                                {
                                    var start = Math.Max(0, foundIdx - 10);
                                    var len = Math.Min(1500, text.Length - start);
                                    report.AppendLine("  【代表性节点片段】");
                                    report.AppendLine(text.Substring(start, len).Trim());
                                    report.AppendLine();
                                }
                            }
                            else
                            {
                                // JSON：输出前 500 字符
                                report.AppendLine(text.Trim().Length > 500 ? text.Trim().Substring(0, 500) + "..." : text.Trim());
                                try
                                {
                                    var json = Newtonsoft.Json.Linq.JToken.Parse(text);
                                    var keys = FindInterestingKeys(json);
                                    if (keys.Count > 0)
                                    {
                                        report.AppendLine();
                                        report.AppendLine($"  相关字段：{string.Join("、", keys.Take(30))}");
                                    }
                                }
                                catch { }
                            }
                            report.AppendLine();
                        }
                        catch { /* 跳过无法读取的条目 */ }
                    }
                }
                else
                {
                    // 纯文本类文件
                    var text = File.ReadAllText(_currentPath, Encoding.UTF8);
                    report.AppendLine("【文件前 500 字符】：");
                    report.AppendLine(text.Length > 500 ? text.Substring(0, 500) + "..." : text);
                }

                // 用只读文本框展示
                var form = new Form
                {
                    Text = "文件内部结构 —— " + Path.GetFileName(_currentPath),
                    Width = 900,
                    Height = 700,
                    StartPosition = FormStartPosition.CenterParent,
                    Font = new Font("Consolas", 9f)
                };
                var txt = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false,
                    Text = report.ToString()
                };
                form.Controls.Add(txt);
                _statusLabel.Text = "已完成结构探查。";
                form.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("探查失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusLabel.Text = "探查失败：" + ex.Message;
            }
        }

        private static int CountWord(string text, string word)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(word)) return 0;
            int count = 0;
            int pos = 0;
            while ((pos = text.IndexOf(word, pos, StringComparison.Ordinal)) != -1)
            {
                count++;
                pos += word.Length;
                if (count > 10000) break; // 防无限循环
            }
            return count;
        }

        private static bool IsZipFile(string path)
        {
            try
            {
                using var fs = File.OpenRead(path);
                var buf = new byte[4];
                if (fs.Read(buf, 0, 4) < 4) return false;
                return buf[0] == 0x50 && buf[1] == 0x4B && (buf[2] == 0x03 || buf[2] == 0x05 || buf[2] == 0x07);
            }
            catch { return false; }
        }

        private static HashSet<string> FindInterestingKeys(Newtonsoft.Json.Linq.JToken token)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var candidates = new[] { "rootTopic","topic","Topic","title","Title","text","children",
                                     "nodes","node","notes","remark","image","images","src","url",
                                     "hyperlink","href","sheet","sheets","mindMap","centerTopic",
                                     "mainTopic","subTopic","rootNode","centralTopic","content",
                                     "doc","document","中心主题","标题","主题","子节点","备注","图片" };
            foreach (var prop in ((Newtonsoft.Json.Linq.JContainer)token).Descendants())
            {
                if (prop is Newtonsoft.Json.Linq.JProperty p)
                {
                    var name = p.Name;
                    if (Array.IndexOf(candidates, name) >= 0) set.Add(name);
                    else
                    {
                        foreach (var c in candidates)
                            if (name.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0)
                                set.Add(name);
                    }
                    if (set.Count > 50) break;
                }
            }
            return set;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
        }
    }
}
