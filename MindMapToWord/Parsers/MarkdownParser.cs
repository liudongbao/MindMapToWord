using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace MindMapToWord.Parsers
{
    using MindMapToWord.Core;

    /// <summary>
    /// Markdown 标题层级 / 列表嵌套 —— 也能作为"导图"输入，便于扩展
    /// </summary>
    public class MarkdownParser : IMindMapParser
    {
        public string FormatName => "Markdown";

        public string[] SupportedExtensions => new[] { ".md", ".markdown", ".txt" };

        public bool CanParse(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            return SupportedExtensions.Contains(ext);
        }

        public MindMapDocument Parse(string filePath)
        {
            var doc = new MindMapDocument
            {
                Title = Path.GetFileNameWithoutExtension(filePath),
                SourceFormat = "Markdown",
                SourcePath = filePath
            };

            var lines = File.ReadAllLines(filePath);

            // 根节点
            var root = new MindMapNode { Title = doc.Title, Level = 0 };
            doc.Roots.Add(root);

            // 用于追踪当前各级最近的父节点（Level => 节点）
            var stack = new Dictionary<int, MindMapNode>();
            stack[0] = root;

            var imageRegex = new Regex(@"!\[([^\]]*)\]\(([^)]+)\)", RegexOptions.Compiled);

            foreach (var rawLine in lines)
            {
                var line = rawLine.TrimEnd();
                if (string.IsNullOrWhiteSpace(line)) continue;

                int level;
                string content;

                // 标题 # 语法
                var headingMatch = Regex.Match(line, @"^(#{1,6})\s+(.+)$");
                if (headingMatch.Success)
                {
                    level = headingMatch.Groups[1].Length;
                    content = headingMatch.Groups[2].Value.Trim();
                }
                else
                {
                    // 列表（* / - / + / 数字.） —— 以缩进数计算层级
                    var indent = line.TakeWhile(char.IsWhiteSpace).Count();
                    var trimmed = line.TrimStart();
                    if (Regex.IsMatch(trimmed, @"^([-*+]|\d+\.)\s+"))
                    {
                        level = (indent / 2) + 1; // 每两个空格代表一层
                        content = Regex.Replace(trimmed, @"^([-*+]|\d+\.)\s+", string.Empty).Trim();
                    }
                    else
                    {
                        // 普通段落挂在当前最深节点的备注
                        var currentNode = stack.Values.OrderByDescending(n => n.Level).FirstOrDefault() ?? root;
                        currentNode.Notes = string.IsNullOrWhiteSpace(currentNode.Notes)
                            ? trimmed
                            : currentNode.Notes + "\n" + trimmed;
                        continue;
                    }
                }

                // 提取图片：在内容中寻找 ![alt](path)
                var images = new List<ImageResource>();
                content = imageRegex.Replace(content, m =>
                {
                    var alt = m.Groups[1].Value.Trim();
                    var src = m.Groups[2].Value.Trim();
                    var img = TryLoadImage(src, filePath);
                    if (img != null)
                    {
                        img.Caption = alt;
                        images.Add(img);
                    }
                    return string.Empty;
                }).Trim();

                // 找到父节点：在 stack 中小于当前 level 的最大 level
                MindMapNode? parent = null;
                foreach (var key in stack.Keys)
                {
                    if (key < level && (parent == null || key > parent.Level)) parent = stack[key];
                }
                parent ??= root;

                var node = new MindMapNode { Title = content, Level = level };
                foreach (var img in images) node.Images.Add(img);
                parent.AddChild(node);

                // 清理 stack 中 >= 当前 level 的条目，再挂入
                foreach (var key in stack.Keys.Where(k => k >= level).ToList()) stack.Remove(key);
                stack[level] = node;
            }

            return doc;
        }

        private static ImageResource? TryLoadImage(string path, string sourcePath)
        {
            try
            {
                if (Uri.IsWellFormedUriString(path, UriKind.Absolute))
                {
                    // 不联网下载，放弃
                    return null;
                }
                var full = Path.IsPathRooted(path) ? path : Path.Combine(Path.GetDirectoryName(sourcePath)!, path);
                if (File.Exists(full))
                    return ImageResource.FromBytes(File.ReadAllBytes(full), Path.GetFileName(full));
            }
            catch { }
            return null;
        }
    }
}
