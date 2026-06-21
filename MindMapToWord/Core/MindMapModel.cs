using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace MindMapToWord.Core
{
    /// <summary>
    /// 思维导图节点模型 —— 统一表示各种导图格式的节点
    /// </summary>
    public class MindMapNode
    {
        public string Title { get; set; } = string.Empty;

        public string? Notes { get; set; }

        public string? Hyperlink { get; set; }

        public List<ImageResource> Images { get; set; } = new List<ImageResource>();

        public List<MindMapNode> Children { get; set; } = new List<MindMapNode>();

        public MindMapNode? Parent { get; set; }

        public int Level { get; set; }

        /// <summary>XML 中的原始顺序索引（调试用）</summary>
        public int RawOrder { get; set; }

        /// <summary>XML 中找到的 Y 坐标值（调试用）</summary>
        public double RawY { get; set; }

        /// <summary>XML 中找到的 X 坐标值（调试用）</summary>
        public double RawX { get; set; }

        /// <summary>XML Shape 的所有属性名=值 拼接（调试用，前300字符）</summary>
        public string? RawAttributes { get; set; }

        /// <summary>XML LevelData 的完整字符串（调试用，前300字符）</summary>
        public string? RawLevelData { get; set; }

        public bool HasContent =>
            !string.IsNullOrWhiteSpace(Title) ||
            !string.IsNullOrWhiteSpace(Notes) ||
            !string.IsNullOrWhiteSpace(Hyperlink) ||
            Images.Count > 0 ||
            Children.Count > 0;

        public void AddChild(MindMapNode child)
        {
            child.Parent = this;
            child.Level = this.Level + 1;
            Children.Add(child);
        }
    }

    /// <summary>
    /// 图片资源 —— 包含二进制数据与格式信息，便于插入 Word
    /// </summary>
    public class ImageResource
    {
        public byte[] Data { get; set; } = Array.Empty<byte>();

        public string FileName { get; set; } = string.Empty;

        public string ContentType { get; set; } = "image/png";

        public int WidthPx { get; set; }

        public int HeightPx { get; set; }

        public string? Caption { get; set; }

        public static ImageResource FromBytes(byte[] data, string fileName, string? caption = null)
        {
            var img = new ImageResource
            {
                Data = data,
                FileName = fileName,
                ContentType = InferContentType(fileName),
                Caption = caption
            };

            try
            {
                using var ms = new MemoryStream(data);
                using var bmp = Image.FromStream(ms);
                img.WidthPx = bmp.Width;
                img.HeightPx = bmp.Height;
            }
            catch
            {
                img.WidthPx = 400;
                img.HeightPx = 300;
            }

            return img;
        }

        private static string InferContentType(string fileName)
        {
            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            return ext switch
            {
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".tif" or ".tiff" => "image/tiff",
                ".webp" => "image/webp",
                _ => "image/png"
            };
        }
    }

    /// <summary>
    /// 思维导图文档 —— 包含若干个根节点（有些导图支持多中心）
    /// </summary>
    public class MindMapDocument
    {
        public string Title { get; set; } = string.Empty;

        public string SourceFormat { get; set; } = string.Empty;

        public string SourcePath { get; set; } = string.Empty;

        public List<MindMapNode> Roots { get; set; } = new List<MindMapNode>();

        public int TotalNodes
        {
            get
            {
                int count = 0;
                foreach (var r in Roots) CountNodes(r, ref count);
                return count;
            }
        }

        public int TotalImages
        {
            get
            {
                int count = 0;
                foreach (var r in Roots) CountImages(r, ref count);
                return count;
            }
        }

        private static void CountNodes(MindMapNode node, ref int count)
        {
            count++;
            foreach (var c in node.Children) CountNodes(c, ref count);
        }

        private static void CountImages(MindMapNode node, ref int count)
        {
            count += node.Images.Count;
            foreach (var c in node.Children) CountImages(c, ref count);
        }
    }
}
