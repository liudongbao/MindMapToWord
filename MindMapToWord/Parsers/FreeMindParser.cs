using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace MindMapToWord.Parsers
{
    using MindMapToWord.Core;

    /// <summary>
    /// FreeMind / MindManager / MindMaster 等基于 .mm (XML) 格式的解析器
    /// </summary>
    public class FreeMindParser : IMindMapParser
    {
        public string FormatName => "FreeMind";

        public string[] SupportedExtensions => new[] { ".mm", ".mmap", ".mindmanager" };

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
                SourceFormat = "FreeMind",
                SourcePath = filePath
            };

            var xdoc = XDocument.Load(filePath);
            var map = xdoc.Root;
            if (map == null) throw new InvalidDataException(".mm 文件缺少根 map 节点");

            var roots = map.Elements("node");
            foreach (var root in roots)
            {
                var node = ParseNode(root, filePath, 0);
                if (node != null && node.HasContent)
                {
                    node.Title = string.IsNullOrWhiteSpace(node.Title) ? doc.Title : node.Title;
                    doc.Roots.Add(node);
                }
            }

            return doc;
        }

        private MindMapNode ParseNode(XElement el, string sourcePath, int level)
        {
            var node = new MindMapNode
            {
                Title = CleanAttr(el.Attribute("TEXT") ?? el.Attribute("text") ?? el.Attribute("TEXT_CDATA") ?? el.Attribute("title")),
                Hyperlink = CleanAttr(el.Attribute("LINK") ?? el.Attribute("link") ?? el.Attribute("URL")),
                Level = level
            };

            // note / richcontent 备注
            var note = el.Elements("note").FirstOrDefault();
            if (note != null) node.Notes = CleanText(note.Value);
            var rich = el.Elements("richcontent").FirstOrDefault();
            if (rich != null) node.Notes = (node.Notes ?? string.Empty) + CleanText(rich.Value);

            // 本地图片路径（FILE 等属性）或内嵌 base64
            foreach (var imgAttr in el.Attributes())
            {
                var name = imgAttr.Name.LocalName.ToUpperInvariant();
                if ((name == "BACKGROUND_IMAGE" || name == "IMAGE" || name == "ICON") &&
                    !string.IsNullOrWhiteSpace(imgAttr.Value))
                {
                    var path = imgAttr.Value;
                    if (File.Exists(path))
                        node.Images.Add(ImageResource.FromBytes(File.ReadAllBytes(path), Path.GetFileName(path)));
                    else if (Uri.IsWellFormedUriString(path, UriKind.Relative))
                    {
                        var full = Path.Combine(Path.GetDirectoryName(sourcePath)!, path);
                        if (File.Exists(full))
                            node.Images.Add(ImageResource.FromBytes(File.ReadAllBytes(full), Path.GetFileName(full)));
                    }
                }
            }

            // 内嵌图标图片（少见）
            foreach (var icon in el.Elements("icon"))
            {
                var builtin = icon.Attribute("BUILTIN")?.Value;
                if (!string.IsNullOrWhiteSpace(builtin) && File.Exists(builtin))
                    node.Images.Add(ImageResource.FromBytes(File.ReadAllBytes(builtin), Path.GetFileName(builtin)));
            }

            // 子节点
            foreach (var child in el.Elements("node"))
            {
                node.AddChild(ParseNode(child, sourcePath, level + 1));
            }

            return node;
        }

        private static string CleanAttr(XAttribute? attr) => attr == null ? string.Empty : CleanText(attr.Value);

        private static string CleanText(string txt)
        {
            if (string.IsNullOrWhiteSpace(txt)) return string.Empty;
            return txt.Trim().Replace("\r\n", "\n").Replace('\r', '\n');
        }
    }
}
