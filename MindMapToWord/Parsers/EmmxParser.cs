using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using MindMapToWord.Core;

namespace MindMapToWord.Parsers
{
    /// <summary>
    /// 亿图脑图 / MindMaster / WPS 脑图（.emmx）解析器。
    /// EMMX 本质为 ZIP 包，本解析器采用"多策略试探"以兼容不同厂商/版本：
    ///   1) 优先查找 document.json / content.json（MindMaster 新版）
    ///   2) 扫描 ZIP 中所有 .json 文件，尝试提取含 topic / nodes / children 信息的根
    ///   3) 扫描所有 .xml 文件，尝试识别 XML 式节点
    ///   4) 若仍失败，将 ZIP 内结构摘要输出到错误消息，便于继续扩充
    ///
    /// 节点字段兼容：
    ///   标题：title / text / TopicTitle / name / 标题 / 中心主题 / topicText / textContent
    ///   备注：notes / note / remark / comment / description / contentText
    ///   链接：href / hyperlink / link / url
    ///   图片：image.src / imageSrc / images / imageUrl / base64Image / image
    ///   子节点：children / topics / subTopics / nodes / childrenNodes
    /// </summary>
    public class EmmxParser : IMindMapParser
    {
        public string FormatName => "EMMX (亿图脑图 / MindMaster / WPS)";

        public string[] SupportedExtensions => new[] { ".emmx" };

        public bool CanParse(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext == ".emmx";
        }

        public MindMapDocument Parse(string filePath)
        {
            var doc = new MindMapDocument
            {
                Title = Path.GetFileNameWithoutExtension(filePath),
                SourceFormat = "EMMX",
                SourcePath = filePath
            };

            using var zip = ZipFile.OpenRead(filePath);

            // 先列出所有 entry 便于诊断
            var allEntries = zip.Entries.ToList();
            var entrySummary = new StringBuilder();
            foreach (var e in allEntries.Take(50))
                entrySummary.AppendLine($"  - {e.FullName}  ({e.Length} bytes)");
            if (allEntries.Count > 50)
                entrySummary.AppendLine($"  ... (省略其余 {allEntries.Count - 50} 个条目)");

            string? sourceDir = null;
            try { sourceDir = Path.GetDirectoryName(filePath); }
            catch { /* ignore */ }

            // ====== 策略 1：直接定位知名 json ======
            var jsonNames = new[]
            {
                "document.json", "Document.json", "content.json", "Content.json",
                "doc.json", "mindmap.json", "MindMap.json", "main.json",
                "data.json", "tree.json", "root.json", "sheet.json"
            };
            foreach (var name in jsonNames)
            {
                var entry = PickEntryIgnoreCase(zip, name);
                if (entry == null) continue;
                try
                {
                    if (TryParseJson(ReadText(entry), doc, zip, sourceDir))
                        return doc;
                }
                catch { /* continue */ }
            }

            // ====== 策略 2：扫描所有 .json ======
            foreach (var entry in allEntries.Where(e =>
                e.Name.EndsWith(".json", StringComparison.OrdinalIgnoreCase)))
            {
                try
                {
                    if (TryParseJson(ReadText(entry), doc, zip, sourceDir))
                        return doc;
                }
                catch { /* continue */ }
            }

            // ====== 策略 3：扫描所有 .xml ======
            // 优先处理 document.xml 和 page/page.xml（MindMaster 格式）
            var xmlPriority = new[] { "document.xml", "page/page.xml", "page.xml", "content.xml" };
            foreach (var name in xmlPriority)
            {
                var entry = PickEntryIgnoreCase(zip, name);
                if (entry != null)
                {
                    try
                    {
                        using var s = entry.Open();
                        var xdoc = XDocument.Load(s);
                        if (TryParseXml(xdoc, doc, zip, sourceDir))
                            return doc;
                    }
                    catch { /* continue to next */ }
                }
            }

            // 再扫描所有其他 .xml 文件
            foreach (var entry in allEntries.Where(e =>
                e.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            {
                try
                {
                    using var s = entry.Open();
                    var xdoc = XDocument.Load(s);
                    if (TryParseXml(xdoc, doc, zip, sourceDir))
                        return doc;
                }
                catch { /* continue */ }
            }

            // ====== 策略 4：若只有一个大文本文件（非json/xml），尝试解析为通用文本树 ======
            // （略，留给未来扩充）

            // ====== 全部失败：给出便于诊断的错误消息 ======
            var msg = $"无法识别的 .emmx 内部结构。ZIP 包含 {allEntries.Count} 个条目：\n"
                    + entrySummary.ToString()
                    + "\n可将此文件发与开发者，以便扩充解析器。";
            throw new InvalidDataException(msg);
        }

        // ====================== JSON 解析 ======================

        private static bool TryParseJson(string json, MindMapDocument doc, ZipArchive zip,
                                           string? sourceDir)
        {
            if (string.IsNullOrWhiteSpace(json)) return false;

            JToken? root = null;
            try { root = JToken.Parse(json); }
            catch { return false; }

            // 统一用"查找含子节点的对象"的方式来发现根
            var candidates = FindRootTopics(root).ToList();
            if (candidates.Count == 0) return false;

            int idx = 0;
            foreach (var c in candidates)
            {
                var node = ParseJsonNode(c, zip, sourceDir, 0);
                if (node == null || !node.HasContent && node.Children.Count == 0) continue;
                if (string.IsNullOrWhiteSpace(node.Title))
                    node.Title = doc.Title + (idx > 0 ? $"（导图 {idx + 1}）" : string.Empty);
                doc.Roots.Add(node);
                idx++;
            }

            return doc.Roots.Count > 0;
        }

        /// <summary>
        /// 在 JSON token 中查找"看起来像思维导图根节点"的对象。
        /// </summary>
        private static IEnumerable<JToken> FindRootTopics(JToken root)
        {
            if (root == null) yield break;

            // 优先通过知名容器路径找到
            // A. { sheets: [ { rootTopic: {...} } ] } 或 { sheets: [ { topic: {...} } ] }
            if (root["sheets"] is JArray sheets)
            {
                foreach (var sh in sheets)
                {
                    var rt = sh["rootTopic"] ?? sh["RootTopic"] ?? sh["root_topic"]
                             ?? sh["topic"] ?? sh["Topic"] ?? sh["root"] ?? sh["Root"]
                             ?? sh["rootNode"] ?? sh["nodes"]?.First ?? sh["centerTopic"]
                             ?? sh["mainTopic"] ?? sh["centralTopic"];
                    if (rt != null && LooksLikeTopic(rt)) yield return rt;
                }
            }

            // B. { rootTopic: {...} } 等直接结构
            string[] directKeys = { "rootTopic", "RootTopic", "root_topic", "topic", "Topic",
                                     "root", "Root", "rootNode", "nodes", "centerTopic",
                                     "centralTopic", "mainTopic", "center", "central",
                                     "mindMap", "MindMap", "tree", "Tree", "diagram",
                                     "中心主题", "根节点", "导图" };
            foreach (var k in directKeys)
            {
                var v = root[k];
                if (v == null) continue;
                if (v is JObject obj && LooksLikeTopic(obj)) yield return obj;
                if (v is JArray arr)
                    foreach (var item in arr)
                        if (LooksLikeTopic(item)) yield return item;
            }

            // C. 根本身就是一个 topic
            if (root is JObject rootObj && LooksLikeTopic(rootObj))
                yield return rootObj;

            // D. 深度扫描：查找任何带有 "children" / "topics" / "title"+"text" 的对象
            if (root is JContainer container)
            {
                foreach (var obj in container.Descendants())
                {
                    if (obj is JObject jo && LooksLikeTopic(jo)) yield return jo;
                }
            }
        }

        private static bool LooksLikeTopic(JToken t)
        {
            if (t == null || t.Type != JTokenType.Object) return false;
            var o = (JObject)t;
            bool hasTitle = o.Property("title") != null || o.Property("Title") != null
                         || o.Property("text") != null || o.Property("Text") != null
                         || o.Property("name") != null || o.Property("Name") != null
                         || o.Property("TopicTitle") != null || o.Property("topicText") != null
                         || o.Property("content") != null;
            bool hasChildren = o.Property("children") != null || o.Property("topics") != null
                            || o.Property("subTopics") != null || o.Property("nodes") != null
                            || o.Property("childrenNodes") != null || o.Property("Topics") != null
                            || o.Property("Children") != null;
            bool hasImage = o.Property("image") != null || o.Property("imageSrc") != null
                         || o.Property("images") != null || o.Property("imgSrc") != null
                         || o.Property("imageUrl") != null || o.Property("base64Image") != null;
            bool hasNote = o.Property("notes") != null || o.Property("note") != null
                        || o.Property("remark") != null || o.Property("description") != null
                        || o.Property("comment") != null;
            return hasTitle || hasChildren || hasImage || hasNote;
        }

        private static MindMapNode? ParseJsonNode(JToken token, ZipArchive zip,
                                                    string? sourceDir, int level)
        {
            if (token == null || token.Type != JTokenType.Object) return null;

            var node = new MindMapNode
            {
                Level = level,
                Title = ReadNodeText(token, new[] { "title", "Title", "text", "Text", "TopicTitle",
                                                    "topicText", "name", "Name", "topic", "content",
                                                    "主题", "标题", "文本", "textContent" })
                         ?? string.Empty,
                Notes = ReadNodeText(token, new[] { "notes", "note", "Notes", "Note", "remark",
                                                     "Remark", "comment", "Comment", "description",
                                                     "Description", "contentText", "备注", "说明" })
                        ?? string.Empty,
                Hyperlink = ReadNodeText(token, new[] { "href", "hyperlink", "Hyperlink", "link",
                                                        "Link", "url", "URL", "Href" })
                          ?? string.Empty
            };

            // 兼容 notes.plain.content
            if (string.IsNullOrWhiteSpace(node.Notes) && token["notes"] is JObject notesObj)
            {
                var inner = ReadNodeText(notesObj, new[] { "plain", "content", "text", "value" })
                         ?? ReadDeepText(notesObj, new[] { "plain", "content" })
                         ?? ReadDeepText(notesObj, new[] { "html", "content" });
                node.Notes = inner ?? string.Empty;
            }

            // 图片
            CollectImagesFromNode(token, zip, sourceDir, node);

            // 子节点
            string[] childContainerKeys =
            {
                "children", "Children", "topics", "Topics", "subTopics", "SubTopics",
                "subtopics", "nodes", "Nodes", "childrenNodes", "childNodes",
                "子节点", "子主题"
            };
            foreach (var key in childContainerKeys)
            {
                var container = token[key];
                if (container == null) continue;

                if (container is JArray arr)
                {
                    foreach (var c in arr)
                    {
                        var child = ParseJsonNode(c, zip, sourceDir, level + 1);
                        if (child != null) node.AddChild(child);
                    }
                }
                else if (container is JObject co)
                {
                    // 可能是 { attached: [...] }, { topics: [...] } 等
                    foreach (var prop in co.Properties())
                    {
                        if (prop.Value is JArray innerArr)
                        {
                            foreach (var c in innerArr)
                            {
                                var child = ParseJsonNode(c, zip, sourceDir, level + 1);
                                if (child != null) node.AddChild(child);
                            }
                        }
                        else if (prop.Value is JObject innerObj)
                        {
                            var child = ParseJsonNode(innerObj, zip, sourceDir, level + 1);
                            if (child != null) node.AddChild(child);
                        }
                    }
                }
            }

            if (!node.HasContent && node.Children.Count == 0) return null;
            return node;
        }

        private static void CollectImagesFromNode(JToken token, ZipArchive zip,
                                                    string? sourceDir, MindMapNode node)
        {
            // 1. "image": { "src": "..." } 或 "image": "path"
            if (token["image"] is JObject img1)
            {
                var src = ReadNodeText(img1, new[] { "src", "Src", "path", "Path", "url",
                                                     "Url", "value", "Value", "href" });
                if (!string.IsNullOrWhiteSpace(src)) node.Images.Add(ExtractImage(zip, src, sourceDir));
            }
            else if (token["image"] is JValue v1 && v1.Type == JTokenType.String && v1.ToString().Length > 0)
            {
                node.Images.Add(ExtractImage(zip, v1.ToString(), sourceDir));
            }

            // 2. "images": [ ... ]
            if (token["images"] is JArray imgs)
            {
                foreach (var i in imgs)
                {
                    var src = ReadNodeText(i, new[] { "src", "Src", "path", "Path", "url", "Url",
                                                     "value", "Value", "href", "imageUrl" })
                            ?? (i.Type == JTokenType.String ? i.ToString() : null);
                    if (!string.IsNullOrWhiteSpace(src)) node.Images.Add(ExtractImage(zip, src, sourceDir));
                }
            }

            // 3. "imageSrc" / "imagePath" / "imageUrl" / "imgSrc"
            string[] directImageKeys =
            {
                "imageSrc", "imgSrc", "imagePath", "imageUrl", "image_url",
                "picUrl", "picPath", "图片", "图片路径", "图片链接"
            };
            foreach (var key in directImageKeys)
            {
                var v = token[key];
                if (v == null || v.Type != JTokenType.String) continue;
                var s = v.ToString();
                if (!string.IsNullOrWhiteSpace(s)) node.Images.Add(ExtractImage(zip, s, sourceDir));
            }

            // 4. base64 图片
            string[] base64Keys = { "base64Image", "imageBase64", "imageData", "base64", "dataURL" };
            foreach (var key in base64Keys)
            {
                var v = token[key];
                if (v == null || v.Type != JTokenType.String) continue;
                var s = v.ToString();
                if (string.IsNullOrWhiteSpace(s)) continue;
                if (s.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)
                    || s.StartsWith("iVBOR") || s.StartsWith("/9j/")  // PNG / JPEG
                    || s.StartsWith("R0lGOD") || s.StartsWith("Qk"))  // GIF / BMP
                {
                    node.Images.Add(ExtractImage(zip, s, sourceDir));
                }
            }

            // 5. 某些版本在 "media" / "resources" / "attachments" 下挂图片链接
            if (token["media"] is JObject || token["resources"] is JObject)
            {
                foreach (var containerKey in new[] { "media", "resources", "attachments", "attach" })
                {
                    var c = token[containerKey];
                    if (c is JArray arr)
                    {
                        foreach (var it in arr)
                        {
                            var src = ReadNodeText(it, new[] { "src", "url", "path", "href", "file" });
                            if (!string.IsNullOrWhiteSpace(src) && IsImagePath(src))
                                node.Images.Add(ExtractImage(zip, src, sourceDir));
                        }
                    }
                }
            }
        }

        private static bool IsImagePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            var ext = Path.GetExtension(path).Trim('.').ToLowerInvariant();
            var imageExts = new HashSet<string> { "png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp", "svg" };
            if (imageExts.Contains(ext)) return true;
            return path.StartsWith("data:image", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ReadNodeText(JToken token, string[] candidateNames)
        {
            foreach (var name in candidateNames)
            {
                var v = token[name];
                if (v == null) continue;

                switch (v.Type)
                {
                    case JTokenType.String:
                        var s = v.ToString();
                        if (!string.IsNullOrWhiteSpace(s)) return s;
                        break;
                    case JTokenType.Integer:
                    case JTokenType.Float:
                        return v.ToString();
                    default:
                        if (v.HasValues)
                        {
                            var inner = v["content"] ?? v["text"] ?? v["plain"] ?? v["value"] ?? v.First;
                            if (inner != null && inner.Type == JTokenType.String)
                            {
                                var s2 = inner.ToString();
                                if (!string.IsNullOrWhiteSpace(s2)) return s2;
                            }
                        }
                        break;
                }
            }
            return null;
        }

        private static string? ReadDeepText(JToken token, string[] nestedKeys)
        {
            JToken? current = token;
            foreach (var k in nestedKeys)
            {
                current = current[k];
                if (current == null) return null;
            }
            if (current.Type == JTokenType.String) return current.ToString();
            return current.ToString();
        }

        // ====================== XML 解析 ======================

        private static bool TryParseXml(XDocument xdoc, MindMapDocument doc,
                                         ZipArchive zip, string? sourceDir)
        {
            if (xdoc?.Root == null) return false;

            // 直接使用 MindMaster 专用解析器（支持 Shape/Topic/Node 格式）
            return TryParseMindMasterPageXml(xdoc, doc, zip, sourceDir);
        }

        // ====================== MindMaster page.xml 专用解析 ======================

        private static readonly HashSet<string> _connectorTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            "MMConnector", "Connector", "Line", "StraightLine", "Bezier", "Curve", "Arrow"
        };

        private static bool TryParseMindMasterPageXml(XDocument xdoc, MindMapDocument doc,
                                                      ZipArchive zip, string? sourceDir)
        {
            if (xdoc.Root == null) return false;

            var allShapes = xdoc.Descendants()
                .Where(e => e.Name.LocalName.Equals("Shape", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (allShapes.Count == 0) return false;

            // 过滤掉连接器 (Connector/MMConnector) —— 这些不是内容节点
            var contentShapes = allShapes.Where(s => !IsConnectorShape(s)).ToList();

            // 构建 Shape ID -> Element 映射
            var shapeById = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in contentShapes)
            {
                var id = GetAttribute(s, "ID");
                if (!string.IsNullOrWhiteSpace(id))
                    shapeById[id] = s;
            }

            // ====== 策略 1：通过 <LevelData><Super V="父节点ID"/></LevelData> 建立父子关系 ======
            // 这是 MindMaster / 亿图脑图最常见的格式
            var parentMapByLevelData = new Dictionary<string, List<XElement>>(StringComparer.OrdinalIgnoreCase);
            var hasLevelData = false;
            foreach (var s in contentShapes)
            {
                var pid = GetLevelDataSuperId(s);
                if (!string.IsNullOrWhiteSpace(pid))
                {
                    hasLevelData = true;
                    if (!parentMapByLevelData.ContainsKey(pid))
                        parentMapByLevelData[pid] = new List<XElement>();
                    parentMapByLevelData[pid].Add(s);
                }
            }

            // ====== 关键修复：按 SubLevel 属性排序 ======
            // SubLevel 是 MindMaster 在父节点的 LevelData 中定义的子节点 ID 列表（分号分隔）
            // 列表中的顺序就是思维导图中从上到下（或从左到右）的显示顺序
            // 对每个父节点的子节点按 SubLevel 中出现的顺序排列
            foreach (var kvp in parentMapByLevelData)
            {
                // 找到父节点，提取其 SubLevel
                string? parentId = kvp.Key;
                if (shapeById.TryGetValue(parentId, out var parentShape))
                {
                    var orderMap = BuildSubLevelOrderMap(parentShape);
                    if (orderMap.Count > 0)
                    {
                        // 使用 SubLevel 顺序排序
                        kvp.Value.Sort((a, b) =>
                        {
                            var idA = GetAttribute(a, "ID");
                            var idB = GetAttribute(b, "ID");
                            var orderA = orderMap.TryGetValue(idA ?? "", out var oA) ? oA : int.MaxValue;
                            var orderB = orderMap.TryGetValue(idB ?? "", out var oB) ? oB : int.MaxValue;
                            return orderA.CompareTo(orderB);
                        });
                    }
                    else
                    {
                        // 没有 SubLevel 时用原始顺序
                        kvp.Value.Sort((a, b) => CompareShapeOrder(a, b));
                    }
                }
                else
                {
                    kvp.Value.Sort((a, b) => CompareShapeOrder(a, b));
                }
            }

            if (hasLevelData)
            {
                // 根节点：没有 Super 或 Super 不存在于 shapeById 中的节点
                var roots = new List<XElement>();
                foreach (var s in contentShapes)
                {
                    var pid = GetLevelDataSuperId(s);
                    if (string.IsNullOrWhiteSpace(pid) || !shapeById.ContainsKey(pid))
                        roots.Add(s);
                }

                if (roots.Count == 0 && contentShapes.Count > 0)
                    roots.Add(contentShapes[0]);

                // 根节点也按 SubLevel 排序
                var rootOrderMap = BuildSubLevelOrderMapFromRoots(roots);
                if (rootOrderMap.Count > 0)
                {
                    roots.Sort((a, b) =>
                    {
                        var idA = GetAttribute(a, "ID");
                        var idB = GetAttribute(b, "ID");
                        var orderA = rootOrderMap.TryGetValue(idA ?? "", out var oA) ? oA : int.MaxValue;
                        var orderB = rootOrderMap.TryGetValue(idB ?? "", out var oB) ? oB : int.MaxValue;
                        return orderA.CompareTo(orderB);
                    });
                }

                int rootOrder = 0;
                foreach (var root in roots)
                {
                    var node = BuildTreeViaMap(root, shapeById, parentMapByLevelData, zip, sourceDir, 0, rootOrder++);
                    if (node != null && (node.HasContent || node.Children.Count > 0))
                        doc.Roots.Add(node);
                }
                if (doc.Roots.Count > 0) return true;
            }

            // ====== 策略 2：直接用 Children/Topics 元素构建树 ======
            var containedIn = new HashSet<XElement>();
            foreach (var s in allShapes)
            {
                foreach (var childEl in s.Elements())
                {
                    if (childEl.Name.LocalName.Equals("Children", StringComparison.OrdinalIgnoreCase) ||
                        childEl.Name.LocalName.Equals("Topics", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var childShape in childEl.Elements()
                            .Where(e => e.Name.LocalName.Equals("Shape", StringComparison.OrdinalIgnoreCase)))
                            containedIn.Add(childShape);
                    }
                }
            }

            var rootShapes = contentShapes.Where(s => !containedIn.Contains(s)).ToList();
            if (rootShapes.Count > 0)
            {
                int order2 = 0;
                foreach (var root in rootShapes)
                {
                    var node = BuildTreeFromShape(root, zip, sourceDir, order2++);
                    if (node != null && (node.HasContent || node.Children.Count > 0))
                        doc.Roots.Add(node);
                }
                if (doc.Roots.Count > 0) return true;
            }

            // ====== 策略 3：通过 ParentID 属性构建关系 ======
            var parentMapByAttr = new Dictionary<string, List<XElement>>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in contentShapes)
            {
                var pid = GetAttribute(s, "ParentID");
                if (!string.IsNullOrWhiteSpace(pid))
                {
                    if (!parentMapByAttr.ContainsKey(pid))
                        parentMapByAttr[pid] = new List<XElement>();
                    parentMapByAttr[pid].Add(s);
                }
            }

            var roots3 = contentShapes.Where(s => string.IsNullOrWhiteSpace(GetAttribute(s, "ParentID"))).ToList();
            if (roots3.Count > 0)
            {
                int order3 = 0;
                foreach (var root in roots3)
                {
                    var node = BuildTreeViaMap(root, shapeById, parentMapByAttr, zip, sourceDir, 0, order3++);
                    if (node != null && (node.HasContent || node.Children.Count > 0))
                        doc.Roots.Add(node);
                }
                if (doc.Roots.Count > 0) return true;
            }

            // ====== 策略 4：所有内容节点按顺序平铺，第一个为根 ======
            if (contentShapes.Count > 0)
            {
                var first = contentShapes[0];
                var rootNode = BuildTreeViaMap(first, shapeById, parentMapByAttr, zip, sourceDir, 0, 0);
                if (rootNode != null)
                    doc.Roots.Add(rootNode);
            }

            return doc.Roots.Count > 0;
        }

        /// <summary>
        /// 从 Shape 中读取 LevelData/Super 的父节点 ID。
        /// 支持格式：
        ///   <LevelData Super="父ID"/>
        ///   <LevelData><Super V="父ID"/></LevelData>
        ///   <ap:LevelData><ap:Super V="父ID"/></ap:LevelData>
        /// </summary>
        private static string GetLevelDataSuperId(XElement shape)
        {
            // 查找 LevelData 子元素（忽略命名空间）
            var levelData = shape.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("LevelData", StringComparison.OrdinalIgnoreCase));
            if (levelData == null) return string.Empty;

            // 格式1：LevelData 本身有 Super 属性
            var superAttr = levelData.Attribute("Super");
            if (superAttr != null && !string.IsNullOrWhiteSpace(superAttr.Value))
                return superAttr.Value.Trim();

            // 格式2：LevelData 下有 <Super V="父ID"/> 子元素
            var superEl = levelData.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("Super", StringComparison.OrdinalIgnoreCase));
            if (superEl != null)
            {
                var v = superEl.Attribute("V") ?? superEl.Attribute("v") ?? superEl.Attribute("Id") ?? superEl.Attribute("Value");
                if (v != null && !string.IsNullOrWhiteSpace(v.Value))
                    return v.Value.Trim();
                // Super 元素内部文本
                if (!string.IsNullOrWhiteSpace(superEl.Value))
                    return superEl.Value.Trim();
            }

            return string.Empty;
        }

        private static MindMapNode BuildTreeFromShape(XElement shape, ZipArchive zip, string? sourceDir, int orderIndex)
        {
            var node = new MindMapNode
            {
                Level = 0,
                Title = GetShapeText(shape),
                Notes = GetShapeNotes(shape),
                Hyperlink = GetShapeHyperlink(shape),
                RawOrder = orderIndex,
                RawY = GetShapeDoubleAny(shape, "PinY", "Y", "LocPinY", "py", "Top") ?? 0,
                RawX = GetShapeDoubleAny(shape, "PinX", "X", "LocPinX", "px", "Left") ?? 0,
                RawAttributes = CaptureAttributesForDebug(shape),
                RawLevelData = CaptureLevelDataForDebug(shape)
            };

            CollectShapeMedia(shape, zip, sourceDir, node);

            int childOrder = 0;
            foreach (var childEl in shape.Elements())
            {
                if (!childEl.Name.LocalName.Equals("Children", StringComparison.OrdinalIgnoreCase) &&
                    !childEl.Name.LocalName.Equals("Topics", StringComparison.OrdinalIgnoreCase) &&
                    !childEl.Name.LocalName.Equals("Items", StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (var childShape in childEl.Elements()
                    .Where(e => e.Name.LocalName.Equals("Shape", StringComparison.OrdinalIgnoreCase)))
                {
                    if (IsConnectorShape(childShape)) continue;
                    var childNode = BuildTreeFromShape(childShape, zip, sourceDir, childOrder++);
                    if (childNode != null)
                        node.AddChild(childNode);
                }
            }

            return node;
        }

        /// <summary>
        /// 通过父节点映射（支持 LevelData/Super 或 ParentID 等任意方式）递归构建树。
        /// </summary>
        private static MindMapNode BuildTreeViaMap(XElement shape,
                                                   Dictionary<string, XElement> byId,
                                                   Dictionary<string, List<XElement>> parentMap,
                                                   ZipArchive zip,
                                                   string? sourceDir, int depth,
                                                   int orderIndex)
        {
            var node = new MindMapNode
            {
                Level = depth,
                Title = GetShapeText(shape),
                Notes = GetShapeNotes(shape),
                Hyperlink = GetShapeHyperlink(shape),
                RawOrder = orderIndex,
                RawY = GetShapeDoubleAny(shape, "PinY", "Y", "LocPinY", "py", "Top") ?? 0,
                RawX = GetShapeDoubleAny(shape, "PinX", "X", "LocPinX", "px", "Left") ?? 0,
                RawAttributes = CaptureAttributesForDebug(shape),
                RawLevelData = CaptureLevelDataForDebug(shape)
            };

            CollectShapeMedia(shape, zip, sourceDir, node);

            var shapeId = GetAttribute(shape, "ID");
            if (!string.IsNullOrWhiteSpace(shapeId) && parentMap.TryGetValue(shapeId, out var children))
            {
                int childOrder = 0;
                foreach (var child in children)
                {
                    if (GetAttribute(child, "ID") == shapeId) continue; // 自引用保护
                    var childNode = BuildTreeViaMap(child, byId, parentMap, zip, sourceDir, depth + 1, childOrder++);
                    if (childNode != null)
                        node.AddChild(childNode);
                }
            }

            return node;
        }

        /// <summary>
        /// 比较两个 Shape 的显示顺序。思维导图通常从上到下排列，Y 坐标越小越靠上。
        /// 优先使用 ZOrder，其次使用 Y 坐标，最后使用 X 坐标。
        /// 坐标可以从 Shape 的直接属性、LevelData 子元素或 Geometric/Bounds 等子元素中获取。
        /// </summary>
        private static int CompareShapeOrder(XElement a, XElement b)
        {
            // 优先使用 ZOrder
            int zA = GetShapeInt(a, "ZOrder");
            int zB = GetShapeInt(b, "ZOrder");
            if (zA != zB) return zA.CompareTo(zB);

            // 其次使用 Y 坐标
            double? yA = GetShapeDoubleAny(a, "PinY", "Y", "LocPinY", "py", "Top", "TopPos", "YPos");
            double? yB = GetShapeDoubleAny(b, "PinY", "Y", "LocPinY", "py", "Top", "TopPos", "YPos");
            if (yA.HasValue && yB.HasValue) return yA.Value.CompareTo(yB.Value);

            // 最后使用 X 坐标
            double? xA = GetShapeDoubleAny(a, "PinX", "X", "LocPinX", "px", "Left", "LeftPos", "XPos");
            double? xB = GetShapeDoubleAny(b, "PinX", "X", "LocPinX", "px", "Left", "LeftPos", "XPos");
            if (xA.HasValue && xB.HasValue) return xA.Value.CompareTo(xB.Value);

            // 后备：按 ID 排序
            return string.Compare(GetAttribute(a, "ID"), GetAttribute(b, "ID"), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 从 Shape 及其子元素中获取第一个存在的整数属性值。
        /// </summary>
        private static int GetShapeInt(XElement shape, string attrName)
        {
            var attr = shape.Attribute(attrName);
            if (attr != null && int.TryParse(attr.Value, out int direct)) return direct;

            var levelData = shape.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("LevelData", StringComparison.OrdinalIgnoreCase));
            if (levelData != null)
            {
                attr = levelData.Attribute(attrName);
                if (attr != null && int.TryParse(attr.Value, out int ld)) return ld;
            }
            return 0;
        }

        /// <summary>
        /// 从 Shape、LevelData 或 Geometric 等子元素中获取第一个存在的浮点属性值。
        /// </summary>
        private static double? GetShapeDoubleAny(XElement shape, params string[] attrNames)
        {
            foreach (var attrName in attrNames)
            {
                // 1. Shape 直接属性
                var attr = shape.Attribute(attrName);
                if (attr != null && double.TryParse(attr.Value, out double direct)) return direct;

                // 2. LevelData 子元素
                var levelData = shape.Elements()
                    .FirstOrDefault(e => e.Name.LocalName.Equals("LevelData", StringComparison.OrdinalIgnoreCase));
                if (levelData != null)
                {
                    attr = levelData.Attribute(attrName);
                    if (attr != null && double.TryParse(attr.Value, out double ld)) return ld;

                    // 3. LevelData 的子元素（如 Geometric、Bounds）
                    foreach (var child in levelData.Elements())
                    {
                        attr = child.Attribute(attrName);
                        if (attr != null && double.TryParse(attr.Value, out double childVal)) return childVal;

                        // 4. 更深层子元素
                        attr = child.Elements()
                            .SelectMany(e => e.Attributes())
                            .FirstOrDefault(a => a.Name.LocalName.Equals(attrName, StringComparison.OrdinalIgnoreCase));
                        if (attr != null && double.TryParse(attr.Value, out double deepVal)) return deepVal;
                    }
                }

                // 5. Shape 的任何后代属性
                attr = shape.Descendants()
                    .SelectMany(e => e.Attributes())
                    .FirstOrDefault(a => a.Name.LocalName.Equals(attrName, StringComparison.OrdinalIgnoreCase));
                if (attr != null && double.TryParse(attr.Value, out double descVal)) return descVal;
            }
            return null;
        }

        /// <summary>
        /// 调试辅助：将 Shape 的所有属性名=值拼接为字符串（限制长度）
        /// </summary>
        private static string? CaptureAttributesForDebug(XElement shape)
        {
            var attrs = shape.Attributes().ToList();
            if (attrs.Count == 0) return null;
            var parts = attrs.Take(15).Select(a => $"{a.Name.LocalName}={a.Value}");
            var result = string.Join(" | ", parts);
            if (result.Length > 280) result = result.Substring(0, 280) + "...";
            return result;
        }

        /// <summary>
        /// 调试辅助：将 LevelData 元素的完整 XML 内容拼接为字符串（限制长度）
        /// </summary>
        private static string? CaptureLevelDataForDebug(XElement shape)
        {
            var levelData = shape.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("LevelData", StringComparison.OrdinalIgnoreCase));
            if (levelData == null) return null;
            try
            {
                var result = levelData.ToString();
                if (result.Length > 280) result = result.Substring(0, 280) + "...";
                return result;
            }
            catch { return null; }
        }

        /// <summary>
        /// 从 Shape 的 LevelData/SubLevel 中提取子节点 ID 的顺序映射。
        /// 返回 ID -> 顺序索引 的字典。
        /// </summary>
        private static Dictionary<string, int> BuildSubLevelOrderMap(XElement parentShape)
        {
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var levelData = parentShape.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("LevelData", StringComparison.OrdinalIgnoreCase));
            if (levelData == null) return result;

            // 找 SubLevel 元素（可能在不同命名空间下）
            XElement? subLevelEl = levelData.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("SubLevel", StringComparison.OrdinalIgnoreCase));
            if (subLevelEl == null) return result;

            // 尝试从属性 V 中获取分号分隔的 ID 列表
            var vAttr = subLevelEl.Attribute("V");
            if (vAttr == null || string.IsNullOrWhiteSpace(vAttr.Value)) return result;

            var ids = vAttr.Value.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < ids.Length; i++)
            {
                var id = ids[i].Trim();
                if (!string.IsNullOrWhiteSpace(id) && !result.ContainsKey(id))
                    result[id] = i;
            }
            return result;
        }

        /// <summary>
        /// 从根节点集合中查找包含 SubLevel 的 LevelData，用于确定根节点之间的顺序。
        /// 通常根节点（如只有一个根）没有 SubLevel，但如果根节点本身被某个隐含父节点引用，
        /// 它的 SubLevel 会在该父节点的 LevelData 中。
        /// 尝试从所有根节点中找到定义了这些根节点顺序的 SubLevel。
        /// </summary>
        private static Dictionary<string, int> BuildSubLevelOrderMapFromRoots(List<XElement> roots)
        {
            var rootIds = roots.Select(r => GetAttribute(r, "ID")).Where(id => !string.IsNullOrWhiteSpace(id)).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (rootIds.Count == 0) return new Dictionary<string, int>();

            // 尝试：直接从根节点的 LevelData/SubLevel 中提取（如果根节点本身有 SubLevel）
            foreach (var root in roots)
            {
                var orderMap = BuildSubLevelOrderMap(root);
                if (orderMap.Count > 1) // 至少要包含多个子节点才有意义
                    return orderMap;
            }

            return new Dictionary<string, int>();
        }

        private static bool IsConnectorShape(XElement shape)
        {
            var type = GetAttribute(shape, "Type");
            if (string.IsNullOrWhiteSpace(type)) return false;
            return _connectorTypes.Contains(type) ||
                   type.Contains("Connector", StringComparison.OrdinalIgnoreCase) ||
                   type.Contains("Line", StringComparison.OrdinalIgnoreCase) ||
                   type.Contains("Arrow", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetShapeText(XElement shape)
        {
            // 1. Text 子元素（最常见）
            var textEl = shape.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("Text", StringComparison.OrdinalIgnoreCase));
            if (textEl != null && !string.IsNullOrWhiteSpace(textEl.Value))
                return textEl.Value.Trim();

            // 2. TextBody 子元素（Edraw/MindMaster 常见）
            var textBody = shape.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("TextBody", StringComparison.OrdinalIgnoreCase));
            if (textBody != null && !string.IsNullOrWhiteSpace(textBody.Value))
                return textBody.Value.Trim();

            // 3. 递归查找任何后代中的 Text / TextBody 元素
            var anyText = shape.Descendants()
                .FirstOrDefault(e =>
                    e.Name.LocalName.Equals("Text", StringComparison.OrdinalIgnoreCase) ||
                    e.Name.LocalName.Equals("TextBody", StringComparison.OrdinalIgnoreCase));
            if (anyText != null && !string.IsNullOrWhiteSpace(anyText.Value))
                return anyText.Value.Trim();

            // 4. Text / Title / Name 属性
            foreach (var attrName in new[] { "Text", "text", "Title", "title", "Name", "name", "Content", "Subject", "TopicText" })
            {
                var attr = shape.Attribute(attrName);
                if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
                    return attr.Value.Trim();
            }

            // 5. 直接取元素文本（叶子节点）
            if (!shape.HasElements && !string.IsNullOrWhiteSpace(shape.Value))
                return shape.Value.Trim();

            // 6. 如果 Shape 里任何地方有非空白文本，拼接起来
            var parts = new List<string>();
            foreach (var d in shape.DescendantNodes())
            {
                if (d is XText xt && !string.IsNullOrWhiteSpace(xt.Value))
                    parts.Add(xt.Value.Trim());
            }
            if (parts.Count > 0)
                return string.Join(" ", parts.Take(5));

            return string.Empty;
        }

        private static string GetShapeNotes(XElement shape)
        {
            foreach (var attrName in new[] { "Notes", "Notes", "Remark", "Description", "Comment", "Memo" })
            {
                var attr = shape.Attribute(attrName);
                if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
                    return attr.Value.Trim();
            }
            var noteEl = shape.Element("Notes");
            if (noteEl != null && !string.IsNullOrWhiteSpace(noteEl.Value))
                return noteEl.Value.Trim();
            var remarkEl = shape.Element("Remark");
            if (remarkEl != null && !string.IsNullOrWhiteSpace(remarkEl.Value))
                return remarkEl.Value.Trim();
            return string.Empty;
        }

        private static string GetShapeHyperlink(XElement shape)
        {
            foreach (var attrName in new[] { "Href", "href", "Link", "URL", "Hyperlink" })
            {
                var attr = shape.Attribute(attrName);
                if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
                    return attr.Value.Trim();
            }
            var linkEl = shape.Element("Hyperlink");
            if (linkEl != null && !string.IsNullOrWhiteSpace(linkEl.Value))
                return linkEl.Value.Trim();
            return string.Empty;
        }

        private static void CollectShapeMedia(XElement shape, ZipArchive zip,
                                              string? sourceDir, MindMapNode node)
        {
            // 1. Image 子元素
            foreach (var imgEl in shape.Elements().Where(e =>
                e.Name.LocalName.Equals("Image", StringComparison.OrdinalIgnoreCase) ||
                e.Name.LocalName.Equals("Picture", StringComparison.OrdinalIgnoreCase) ||
                e.Name.LocalName.Equals("Media", StringComparison.OrdinalIgnoreCase)))
            {
                // 从属性中获取路径
                foreach (var attrName in new[] { "Src", "src", "Path", "path", "Ref", "ref",
                                                  "File", "file", "Image", "image" })
                {
                    var attr = imgEl.Attribute(attrName);
                    if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
                    {
                        node.Images.Add(ExtractImage(zip, attr.Value, sourceDir));
                        return;
                    }
                }
                // 也可能在 Image 元素的 Value 中
                if (!string.IsNullOrWhiteSpace(imgEl.Value))
                {
                    node.Images.Add(ExtractImage(zip, imgEl.Value, sourceDir));
                }
            }

            // 2. Shape 上的 Image 属性（某些格式）
            foreach (var attrName in new[] { "Image", "ImagePath", "ImageSrc", "Pic", "Picture" })
            {
                var attr = shape.Attribute(attrName);
                if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
                {
                    node.Images.Add(ExtractImage(zip, attr.Value, sourceDir));
                    return;
                }
            }

            // 3. media/ 相对路径（MindMaster 通常将图片放在 media/ 目录）
            // 检查是否引用了 media/img*.png 等文件
            foreach (var attr in shape.Attributes())
            {
                var val = attr.Value;
                if (!string.IsNullOrWhiteSpace(val) &&
                    (val.Contains("media/", StringComparison.OrdinalIgnoreCase) ||
                     val.Contains(".png", StringComparison.OrdinalIgnoreCase) ||
                     val.Contains(".jpg", StringComparison.OrdinalIgnoreCase) ||
                     val.Contains(".gif", StringComparison.OrdinalIgnoreCase)))
                {
                    node.Images.Add(ExtractImage(zip, val, sourceDir));
                }
            }
        }

        private static string GetAttribute(XElement e, string name)
        {
            var attr = e.Attribute(name);
            return attr?.Value?.Trim() ?? string.Empty;
        }

        // ====================== 工具方法 ======================

        private static ZipArchiveEntry? PickEntryIgnoreCase(ZipArchive zip, string name)
        {
            // 先尝试精确路径
            var e = zip.GetEntry(name)
                   ?? zip.Entries.FirstOrDefault(x =>
                       x.FullName.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (e != null) return e;
            // 再按文件名匹配
            var fileName = Path.GetFileName(name);
            return zip.Entries.FirstOrDefault(x =>
                x.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase));
        }

        private static string ReadText(ZipArchiveEntry entry)
        {
            using var s = entry.Open();
            using var sr = new StreamReader(s, Encoding.UTF8);
            return sr.ReadToEnd();
        }

        private static ImageResource ExtractImage(ZipArchive zip, string source, string? sourceDir)
        {
            if (string.IsNullOrWhiteSpace(source))
                return EmptyImage("空路径");

            // A. base64 内联图片
            if (source.StartsWith("data:image", StringComparison.OrdinalIgnoreCase))
            {
                var commaIdx = source.IndexOf(',');
                var mime = commaIdx > 0 ? source.Substring(5, commaIdx - 5) : "image/png";
                var b64 = commaIdx > 0 ? source.Substring(commaIdx + 1) : source;
                try
                {
                    var data = Convert.FromBase64String(b64);
                    var ext = mime.Split('/')[1].Split(';')[0];
                    return ImageResource.FromBytes(data, "inline." + ext);
                }
                catch { /* fall through */ }
            }
            if (source.StartsWith("iVBOR") || source.StartsWith("/9j/") ||
                source.StartsWith("R0lGOD") || source.StartsWith("Qk"))
            {
                try
                {
                    var data = Convert.FromBase64String(source);
                    return ImageResource.FromBytes(data, "inline.png");
                }
                catch { /* fall through */ }
            }

            // B. http(s) 链接 —— 不联网下载，仅当文字链接
            if (source.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                return new ImageResource
                {
                    FileName = source,
                    Data = Array.Empty<byte>(),
                    Caption = "网络图片（未下载）：" + source,
                    ContentType = "image/png",
                    WidthPx = 0,
                    HeightPx = 0
                };

            // C. ZIP 包内路径
            var cleanPath = source.Trim().TrimStart('/').Replace('\\', '/');
            // 去掉可能的 "./" 前缀
            if (cleanPath.StartsWith("./", StringComparison.Ordinal))
                cleanPath = cleanPath.Substring(2);

            var entry = zip.GetEntry(cleanPath)
                       ?? zip.Entries.FirstOrDefault(e =>
                           e.FullName.Replace('\\', '/').Equals(cleanPath, StringComparison.OrdinalIgnoreCase))
                       ?? zip.Entries.FirstOrDefault(e =>
                           e.Name.Equals(Path.GetFileName(cleanPath), StringComparison.OrdinalIgnoreCase));
            if (entry != null)
            {
                using var ms = new MemoryStream();
                entry.Open().CopyTo(ms);
                return ImageResource.FromBytes(ms.ToArray(), entry.Name);
            }

            // D. 本地绝对路径
            if (File.Exists(source))
                return ImageResource.FromBytes(File.ReadAllBytes(source), Path.GetFileName(source));

            // E. 相对于源文件所在目录
            if (!string.IsNullOrWhiteSpace(sourceDir) && Directory.Exists(sourceDir))
            {
                var local = Path.Combine(sourceDir, cleanPath);
                if (File.Exists(local))
                    return ImageResource.FromBytes(File.ReadAllBytes(local), Path.GetFileName(local));

                // 仅按文件名匹配（有些版本只写文件名）
                var byName = Path.Combine(sourceDir, Path.GetFileName(cleanPath));
                if (File.Exists(byName))
                    return ImageResource.FromBytes(File.ReadAllBytes(byName), Path.GetFileName(byName));
            }

            return new ImageResource
            {
                FileName = Path.GetFileName(cleanPath),
                Data = Array.Empty<byte>(),
                ContentType = "image/png",
                WidthPx = 0,
                HeightPx = 0,
                Caption = "图片未找到：" + cleanPath
            };
        }

        private static ImageResource EmptyImage(string caption)
            => new() { FileName = "empty", Data = Array.Empty<byte>(),
                       ContentType = "image/png", WidthPx = 0, HeightPx = 0, Caption = caption };
    }
}
