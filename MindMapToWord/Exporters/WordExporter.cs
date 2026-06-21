using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WPColor = DocumentFormat.OpenXml.Wordprocessing.Color;
using MindMapToWord.Core;

namespace MindMapToWord.Exporters
{
    /// <summary>
    /// 将 MindMapDocument 写入 .docx。使用 OpenXML SDK，不依赖 Office。
    /// </summary>
    public class WordExporter
    {
        public string OutputPath { get; }
        public bool IncludeNotes { get; set; } = true;
        public bool IncludeHyperlinks { get; set; } = true;

        public WordExporter(string outputPath)
        {
            OutputPath = outputPath;
        }

        public void Export(MindMapDocument doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (doc.Roots.Count == 0) throw new InvalidDataException("思维导图为空，无可导出内容。");

            using var package = WordprocessingDocument.Create(OutputPath, WordprocessingDocumentType.Document);
            var mainPart = package.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // 文档属性
            SetDocProperties(package, mainPart, doc);

            // 添加样式
            AddStyleDefinitions(mainPart);

            // 标题
            AppendTitle(body, doc.Title);

            // 导图层级内容
            foreach (var root in doc.Roots)
            {
                AppendNode(body, mainPart, root, 1);
            }

            mainPart.Document.Save();
        }

        private static void SetDocProperties(
            WordprocessingDocument package,
            MainDocumentPart main,
            MindMapDocument doc)
        {
            // 核心属性（标题）
            var corePart = package.AddCoreFilePropertiesPart();
            using var coreWriter = new StreamWriter(corePart.GetStream(FileMode.Create));
            coreWriter.Write(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                "<dc:title>" + EscapeXml(doc.Title) + "</dc:title>" +
                "<dc:creator>MindMapToWord</dc:creator>" +
                "<cp:lastModifiedBy>MindMapToWord</cp:lastModifiedBy>" +
                "</cp:coreProperties>");

            // 扩展属性（应用程序标识）
            var extPart = package.AddExtendedFilePropertiesPart();
            using var extWriter = new StreamWriter(extPart.GetStream(FileMode.Create));
            extWriter.Write(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>MindMapToWord</Application>" +
                "<Company>MindMapToWord</Company>" +
                "</Properties>");
        }

        private static string EscapeXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        private static void AddStyleDefinitions(MainDocumentPart main)
        {
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();

            var styles = new Styles();

            // ===== 默认段落样式 Normal =====
            var normal = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                CustomStyle = true
            };
            normal.Append(new StyleName { Val = "Normal" });
            normal.Append(new PrimaryStyle { Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On) });
            normal.Append(new StyleRunProperties(
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "微软雅黑" },
                new FontSize { Val = "22" }
            ));
            styles.Append(normal);

            // ===== 标题 1~4 —— 使用 Word 内置样式 ID 和 style name，确保导航窗格可识别 =====
            // 关键：样式名使用 Word 内部命名 "heading 1" / "heading 2"（半角小写 + 空格），
            // 样式 ID 使用 "Heading1" / "Heading2"（无空格，首字母大写）。
            styles.Append(CreateHeadingStyle("Heading1", "heading 1", 36, "1F4E79", 0));
            styles.Append(CreateHeadingStyle("Heading2", "heading 2", 28, "2E75B6", 1));
            styles.Append(CreateHeadingStyle("Heading3", "heading 3", 24, "2E75B6", 2));
            styles.Append(CreateHeadingStyle("Heading4", "heading 4", 22, "404040", 3));

            // ===== 超链接字符样式 =====
            var hyperlink = new Style
            {
                Type = StyleValues.Character,
                StyleId = "Hyperlink",
                CustomStyle = true
            };
            hyperlink.Append(new StyleName { Val = "Hyperlink" });
            hyperlink.Append(new PrimaryStyle { Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On) });
            hyperlink.Append(new StyleRunProperties(
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "微软雅黑" },
                new WPColor { Val = "0563C1" },
                new Underline { Val = UnderlineValues.Single }
            ));
            styles.Append(hyperlink);

            stylesPart.Styles = styles;
        }

        /// <summary>
        /// 创建一个标题样式。outlineLevelValue 为 Word 的大纲级别（0=1级，1=2级，2=3级，3=4级）。
        /// 在段落属性中设置 OutlineLevel 是 Word 导航窗格识别层级的关键。
        /// </summary>
        private static Style CreateHeadingStyle(string id, string name, int sizeHalfPt, string colorHex, int outlineLevelValue)
        {
            var style = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = id,
                CustomStyle = true
            };
            style.Append(new StyleName { Val = name });
            style.Append(new PrimaryStyle { Val = new EnumValue<OnOffOnlyValues>(OnOffOnlyValues.On) });

            // 段落属性：设置大纲级别 + 前后间距
            var stylePPr = new StyleParagraphProperties();
            stylePPr.Append(new OutlineLevel { Val = outlineLevelValue });
            stylePPr.Append(new SpacingBetweenLines { Before = "240", After = "120" });
            style.Append(stylePPr);

            // 字体属性
            var styleRPr = new StyleRunProperties();
            styleRPr.Append(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "微软雅黑" });
            styleRPr.Append(new Bold());
            styleRPr.Append(new FontSize { Val = sizeHalfPt.ToString() });
            styleRPr.Append(new WPColor { Val = colorHex });
            style.Append(styleRPr);
            return style;
        }

        private static void AppendTitle(Body body, string title)
        {
            var para = new Paragraph(new ParagraphProperties(
                new ParagraphStyleId { Val = "Heading1" },
                new OutlineLevel { Val = 0 },
                new Justification { Val = JustificationValues.Center }
            ));
            para.Append(new Run(new RunProperties(new Bold(), new FontSize { Val = "48" }), new Text(title)));
            body.Append(para);
            body.Append(new Paragraph(new Run(new Text(string.Empty))));
        }

        private void AppendNode(Body body, MainDocumentPart main, MindMapNode node, int headingLevel)
        {
            // ====== 叶子节点 + 长文本 → 作为正文段落渲染（不进入目录） ======
            // 判断条件：
            //   1. 没有子节点（叶子节点）
            //   2. 标题文字较长（> 50 字符）—— 说明是正文段落而非章节标题
            //   3. 层级较深（>= 第 3 级）—— 浅层级的节点即使长文本也作为标题
            // 这类节点本质是正文内容，不设大纲级别，不出现在 Word 导航窗格中
            bool isLeafLongText = node.Children.Count == 0
                                   && (node.Title ?? string.Empty).Length > 50
                                   && headingLevel >= 3;

            if (isLeafLongText)
            {
                // 正文段落：使用 Normal 样式，不加 outlineLvl，不加粗
                var bodyPara = new Paragraph();
                bodyPara.Append(new Run(new Text(string.IsNullOrWhiteSpace(node.Title) ? "（无标题）" : node.Title)));
                body.Append(bodyPara);

                // 图片（正文段落也可能配图）
                AppendImages(body, main, node);
                return;
            }

            // ====== 标题段落渲染 ======
            var clampedLevel = Math.Min(Math.Max(headingLevel, 1), 4);
            var styleId = "Heading" + clampedLevel;
            // Word 大纲级别：0=级别1，1=级别2，2=级别3，3=级别4
            var outlineLevel = clampedLevel - 1;

            // 同时设置：样式 ID + 直接设置 outlineLvl（双重保险，导航窗格依赖 outlineLvl）
            var pPr = new ParagraphProperties(
                new ParagraphStyleId { Val = styleId },
                new OutlineLevel { Val = outlineLevel }
            );
            var titlePara = new Paragraph(pPr);
            titlePara.Append(new Run(new Text(string.IsNullOrWhiteSpace(node.Title) ? "（无标题）" : node.Title)));
            body.Append(titlePara);

            // 超链接
            if (IncludeHyperlinks && !string.IsNullOrWhiteSpace(node.Hyperlink))
            {
                var hyperlinkRel = main.AddHyperlinkRelationship(new Uri(node.Hyperlink), true);
                var hyperlink = new Hyperlink(
                    new Run(
                        new RunProperties(
                            new RunStyle { Val = "Hyperlink" },
                            new WPColor { Val = "0563C1" },
                            new Underline { Val = UnderlineValues.Single }
                        ),
                        new Text(node.Hyperlink)
                    )
                )
                { History = new OnOffValue(true), Id = hyperlinkRel.Id };

                var linkPara = new Paragraph();
                linkPara.Append(new Run(new Text("链接：")));
                linkPara.Append(hyperlink);
                body.Append(linkPara);
            }

            // 备注
            if (IncludeNotes && !string.IsNullOrWhiteSpace(node.Notes))
            {
                foreach (var line in node.Notes.Split('\n', StringSplitOptions.None))
                {
                    var p = new Paragraph();
                    p.Append(new Run(new RunProperties(new Italic(), new FontSize { Val = "20" }), new Text(line)));
                    body.Append(p);
                }
            }

            // 图片
            AppendImages(body, main, node);

            // 子节点
            foreach (var child in node.Children)
            {
                AppendNode(body, main, child, headingLevel + 1);
            }
        }

        /// <summary>
        /// 渲染节点附带的图片（与节点标题在同一层级的独立段落）
        /// </summary>
        private void AppendImages(Body body, MainDocumentPart main, MindMapNode node)
        {
            foreach (var img in node.Images)
            {
                if (img.Data == null || img.Data.Length == 0) continue;
                try
                {
                    var imagePartId = AddImagePart(main, img);
                    var picture = CreatePicture(imagePartId, img);
                    var imgPara = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
                    imgPara.Append(new Run(picture));
                    body.Append(imgPara);

                    if (!string.IsNullOrWhiteSpace(img.Caption))
                    {
                        var cap = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
                        cap.Append(new Run(new RunProperties(new Italic(), new FontSize { Val = "18" }), new Text("图：" + img.Caption)));
                        body.Append(cap);
                    }
                }
                catch
                {
                    var p = new Paragraph();
                    p.Append(new Run(new RunProperties(new Italic()), new Text("[图片 '" + img.FileName + "' 无法显示]")));
                    body.Append(p);
                }
            }
        }

        private static string AddImagePart(MainDocumentPart main, ImageResource img)
        {
            ImagePartType type = img.ContentType.ToLowerInvariant() switch
            {
                "image/jpeg" => ImagePartType.Jpeg,
                "image/gif" => ImagePartType.Gif,
                "image/bmp" => ImagePartType.Bmp,
                "image/tiff" => ImagePartType.Tiff,
                _ => ImagePartType.Png
            };

            var part = main.AddImagePart(type);
            using var ms = new MemoryStream(img.Data);
            part.FeedData(ms);
            return main.GetIdOfPart(part);
        }

        private static Drawing CreatePicture(string relationshipId, ImageResource img)
        {
            // EMU：96 DPI → 像素 * 9525
            const long emuPerPixel = 9525L;
            long maxWidthEmu = 6L * 914400L; // 6 英寸上限
            long widthEmu = Math.Min(img.WidthPx * emuPerPixel, maxWidthEmu);
            long heightEmu = img.WidthPx > 0
                ? (long)(widthEmu * ((double)img.HeightPx / img.WidthPx))
                : img.HeightPx * emuPerPixel;

            var inline = new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = (UInt32Value)1U, Name = "Picture" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = img.FileName },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U
            };

            return new Drawing(inline);
        }
    }
}
