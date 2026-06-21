using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;

namespace MindMapToWord.Parsers
{
    using MindMapToWord.Core;

    /// <summary>
    /// XMind (.xmind) 解析器。XMind 文件本质是 ZIP 包，内部包含 content.json / content.xml。
    /// 同时兼容 XMind 2020+ 与 XMind 8/Zen 两种版本。
    /// </summary>
    public class XMindParser : IMindMapParser
    {
        public string FormatName => "XMind";

        public string[] SupportedExtensions => new[] { ".xmind" };

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
                SourceFormat = "XMind",
                SourcePath = filePath
            };

            using var zip = ZipFile.OpenRead(filePath);

            // 1) 尝试 JSON 格式（XMind 2020+）
            var jsonEntry = zip.GetEntry("content.json");
            if (jsonEntry != null)
            {
                using var s = jsonEntry.Open();
                using var sr = new StreamReader(s, Encoding.UTF8);
                var json = sr.ReadToEnd();
                ParseJsonSheet(json, doc, zip);
                return doc;
            }

            // 2) 回退到 XML 格式（XMind 8 / Zen）
            var xmlEntry = zip.GetEntry("content.xml");
            if (xmlEntry != null)
            {
                using var s = xmlEntry.Open();
                var xdoc = XDocument.Load(s);
                ParseXmlSheet(xdoc, doc, zip);
                return doc;
            }

            throw new InvalidDataException("无法识别的 XMind 文件结构：缺少 content.json 或 content.xml");
        }

        private void ParseJsonSheet(string json, MindMapDocument doc, ZipArchive zip)
        {
            JArray? sheets = null;
            try
            {
                var token = JToken.Parse(json);
                if (token is JArray arr) sheets = arr;
                else if (token is JObject obj && obj["sheets"] is JArray s2) sheets = s2;
                else if (token is JObject obj2 && obj2["rootTopic"] != null) sheets = new JArray { obj2 };
            }
            catch
            {
                throw new InvalidDataException("XMind content.json 格式错误");
            }

            if (sheets == null) return;

            foreach (var sheet in sheets)
            {
                var rootTopic = sheet["rootTopic"] ?? sheet["topic"];
                if (rootTopic == null) continue;

                var title = (string?)sheet["title"] ?? doc.Title;
                var rootNode = ParseJsonNode(rootTopic, zip, 0);
                if (rootNode != null)
                {
                    rootNode.Title = string.IsNullOrWhiteSpace(rootNode.Title) ? title : rootNode.Title;
                    doc.Roots.Add(rootNode);
                }
            }
        }

        private MindMapNode? ParseJsonNode(JToken token, ZipArchive zip, int level)
        {
            var node = new MindMapNode
            {
                Title = CleanText((string?)token["title"]),
                Notes = CleanText((string?)token["notes"]?["plain"]?["content"] ?? (string?)token["notes"]),
                Hyperlink = (string?)token["href"] ?? (string?)token["link"],
                Level = level
            };

            // 图片 —— JSON 格式中 image/imageSrc 字段通常指向 zip 内路径
            var images = new List<string>();
            if (token["image"] is JObject img && (string?)img["src"] != null)
                images.Add(((string?)img["src"])!.TrimStart('/'));
            if (token["images"] is JArray imgs)
            {
                foreach (var i in imgs)
                {
                    var src = (string?)i["src"] ?? (string?)i;
                    if (!string.IsNullOrWhiteSpace(src)) images.Add(src.TrimStart('/'));
                }
            }
            foreach (var p in images) node.Images.Add(ExtractImage(zip, p));

            // 子节点：children.attached / topics
            JToken? children = null;
            if (token["children"] is JObject chObj)
            {
                children = chObj["attached"] ?? chObj["topics"];
            }
            else if (token["topics"] != null)
            {
                children = token["topics"];
            }

            if (children is JArray arr)
            {
                foreach (var c in arr)
                {
                    var child = ParseJsonNode(c, zip, level + 1);
                    if (child != null) node.AddChild(child);
                }
            }

            return node;
        }

        private void ParseXmlSheet(XDocument xdoc, MindMapDocument doc, ZipArchive zip)
        {
            var ns = xdoc.Root?.GetDefaultNamespace();
            var sheets = xdoc.Descendants((ns ?? XNamespace.None) + "sheet");
            foreach (var sheet in sheets)
            {
                var title = (string?)sheet.Attribute("title");
                var topicName = (ns ?? XNamespace.None) + "topic";
                var rootTopic = sheet.Element(topicName) ?? sheet.Descendants(topicName).FirstOrDefault();
                if (rootTopic == null) continue;

                var rootNode = ParseXmlNode(rootTopic, zip, ns ?? XNamespace.None, 0);
                if (rootNode != null)
                {
                    rootNode.Title = string.IsNullOrWhiteSpace(rootNode.Title) ? (title ?? doc.Title) : rootNode.Title;
                    doc.Roots.Add(rootNode);
                }
            }
        }

        private MindMapNode ParseXmlNode(XElement el, ZipArchive zip, XNamespace ns, int level)
        {
            var node = new MindMapNode
            {
                Title = CleanText((string?)el.Attribute("title") ?? (string?)el.Element(ns + "title")),
                Level = level,
                Hyperlink = (string?)el.Attribute("href")
            };

            var notesEl = el.Descendants(ns + "notes").FirstOrDefault();
            if (notesEl != null) node.Notes = CleanText(notesEl.Value);

            // images
            foreach (var img in el.Descendants(ns + "image"))
            {
                var href = (string?)img.Attribute("href") ?? (string?)img.Attribute("src");
                if (!string.IsNullOrWhiteSpace(href))
                    node.Images.Add(ExtractImage(zip, href.TrimStart('/')));
            }
            // html img tags embedded
            var xhtmlNs = XNamespace.Get("http://www.w3.org/1999/xhtml");
            foreach (var img in el.Descendants(xhtmlNs + "img"))
            {
                var href = (string?)img.Attribute("src");
                if (!string.IsNullOrWhiteSpace(href))
                    node.Images.Add(ExtractImage(zip, href.TrimStart('/')));
            }

            foreach (var child in el.Elements(ns + "topic"))
            {
                node.AddChild(ParseXmlNode(child, zip, ns, level + 1));
            }
            // some formats use topics wrapper
            foreach (var child in el.Elements(ns + "topics").Elements(ns + "topic"))
            {
                node.AddChild(ParseXmlNode(child, zip, ns, level + 1));
            }

            return node;
        }

        private static ImageResource ExtractImage(ZipArchive zip, string entryPath)
        {
            var entry = zip.GetEntry(entryPath)
                        ?? zip.Entries.FirstOrDefault(e => e.FullName.Replace('\\', '/').EndsWith(entryPath, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                return new ImageResource
                {
                    FileName = Path.GetFileName(entryPath),
                    Data = Array.Empty<byte>(),
                    ContentType = "image/png",
                    WidthPx = 0,
                    HeightPx = 0,
                    Caption = "图片未找到：" + entryPath
                };
            }
            using var ms = new MemoryStream();
            entry.Open().CopyTo(ms);
            return ImageResource.FromBytes(ms.ToArray(), entry.Name);
        }

        private static string CleanText(string? txt)
        {
            if (string.IsNullOrWhiteSpace(txt)) return string.Empty;
            return txt.Trim().Replace("\r\n", "\n").Replace('\r', '\n');
        }
    }
}
