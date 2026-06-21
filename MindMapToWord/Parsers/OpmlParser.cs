using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace MindMapToWord.Parsers
{
    using MindMapToWord.Core;

    /// <summary>
    /// OPML (Outline Processor Markup Language) 解析器 —— WorkFlowy、OmniOutliner 等常用
    /// </summary>
    public class OpmlParser : IMindMapParser
    {
        public string FormatName => "OPML";

        public string[] SupportedExtensions => new[] { ".opml", ".xml" };

        public bool CanParse(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".opml" && ext != ".xml") return false;
            try
            {
                var xdoc = XDocument.Load(filePath);
                return xdoc.Root?.Name.LocalName.Equals("opml", StringComparison.OrdinalIgnoreCase) == true;
            }
            catch { return false; }
        }

        public MindMapDocument Parse(string filePath)
        {
            var doc = new MindMapDocument
            {
                Title = Path.GetFileNameWithoutExtension(filePath),
                SourceFormat = "OPML",
                SourcePath = filePath
            };

            var xdoc = XDocument.Load(filePath);
            var headTitle = xdoc.Descendants("head").Elements("title").FirstOrDefault()?.Value;
            if (!string.IsNullOrWhiteSpace(headTitle)) doc.Title = headTitle.Trim();

            var bodies = xdoc.Descendants("body");
            foreach (var body in bodies)
            {
                foreach (var outline in body.Elements("outline"))
                {
                    var node = ParseNode(outline, filePath, 0);
                    if (node != null && node.HasContent) doc.Roots.Add(node);
                }
            }

            return doc;
        }

        private MindMapNode ParseNode(XElement el, string sourcePath, int level)
        {
            var text = el.Attribute("text")?.Value ?? el.Attribute("Text")?.Value ?? el.Attribute("title")?.Value ?? string.Empty;
            var note = el.Attribute("_note")?.Value ?? el.Attribute("note")?.Value;

            var node = new MindMapNode
            {
                Title = CleanText(text),
                Notes = string.IsNullOrWhiteSpace(note) ? null : CleanText(note),
                Hyperlink = el.Attribute("url")?.Value ?? el.Attribute("URL")?.Value ?? el.Attribute("link")?.Value,
                Level = level
            };

            // 图片：查找本地文件路径或 base64
            var imgAttr = el.Attribute("image") ?? el.Attribute("src");
            if (imgAttr != null)
            {
                var p = imgAttr.Value;
                if (File.Exists(p))
                    node.Images.Add(ImageResource.FromBytes(File.ReadAllBytes(p), Path.GetFileName(p)));
                else
                {
                    var full = Path.Combine(Path.GetDirectoryName(sourcePath)!, p);
                    if (File.Exists(full))
                        node.Images.Add(ImageResource.FromBytes(File.ReadAllBytes(full), Path.GetFileName(full)));
                }
            }

            foreach (var child in el.Elements("outline"))
            {
                node.AddChild(ParseNode(child, sourcePath, level + 1));
            }

            return node;
        }

        private static string CleanText(string txt)
        {
            if (string.IsNullOrWhiteSpace(txt)) return string.Empty;
            return txt.Trim().Replace("\r\n", "\n").Replace('\r', '\n');
        }
    }
}
