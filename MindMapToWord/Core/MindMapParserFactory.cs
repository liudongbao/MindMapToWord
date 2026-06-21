using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MindMapToWord.Core
{
    /// <summary>
    /// 解析器工厂 —— 根据文件扩展名自动匹配最合适的解析器
    /// </summary>
    public static class MindMapParserFactory
    {
        private static readonly List<IMindMapParser> Parsers = new()
        {
            new Parsers.EmmxParser(),
            new Parsers.XMindParser(),
            new Parsers.FreeMindParser(),
            new Parsers.OpmlParser(),
            new Parsers.MarkdownParser()
        };

        public static IReadOnlyList<IMindMapParser> All => Parsers;

        public static string SupportedFilter => "思维导图文件|*.xmind;*.emmx;*.mm;*.mmap;*.opml;*.md;*.markdown;*.txt|所有文件|*.*";

        public static IMindMapParser Select(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();

            // EMMX 优先 —— ZIP 包
            var emmx = Parsers.OfType<Parsers.EmmxParser>().First();
            if (ext == ".emmx") return emmx;

            // XMind
            var xmind = Parsers.OfType<Parsers.XMindParser>().First();
            if (xmind.CanParse(filePath)) return xmind;

            // OPML 有自己的校验逻辑
            var opml = Parsers.OfType<Parsers.OpmlParser>().First();
            if (ext == ".opml" || ext == ".xml")
            {
                try { if (opml.CanParse(filePath)) return opml; } catch { }
            }

            // 按扩展名匹配
            foreach (var p in Parsers)
            {
                if (p.SupportedExtensions.Contains(ext)) return p;
            }

            // 兜底：尝试 Markdown
            return Parsers.OfType<Parsers.MarkdownParser>().First();
        }

        public static MindMapDocument Parse(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("文件不存在：" + filePath);

            var parser = Select(filePath);
            return parser.Parse(filePath);
        }
    }
}
