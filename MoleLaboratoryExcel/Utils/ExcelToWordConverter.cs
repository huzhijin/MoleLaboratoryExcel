using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
// 添加图片处理所需的命名空间
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

public class ExcelToWordConverter
{
    private readonly DataFormatter _formatter;

    public ExcelToWordConverter()
    {
        // 创建DataFormatter实例，用于格式化单元格值
        _formatter = new DataFormatter(true); // true表示使用本地化格式
    }

    public void ConvertExcelToWord(string excelPath, string wordPath)
    {
        // 检查文件路径
        if (string.IsNullOrEmpty(excelPath))
            throw new ArgumentException("Excel文件路径不能为空", nameof(excelPath));
        if (string.IsNullOrEmpty(wordPath))
            throw new ArgumentException("Word文件路径不能为空", nameof(wordPath));
        if (!File.Exists(excelPath))
            throw new FileNotFoundException("找不到Excel文件", excelPath);

        try
        {
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fs);

                using (WordprocessingDocument wordDoc =
                    WordprocessingDocument.Create(wordPath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // 添加目录标题
                    AddChapterTitle(body, "目录", 1, true);

                    // 创建目录段落
                    var tocPara = body.AppendChild(new Paragraph());

                    // 创建目录字段
                    var run = tocPara.AppendChild(new Run());
                    run.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Begin });

                    var run2 = tocPara.AppendChild(new Run());
                    run2.AppendChild(new Text(" TOC \\h \\z "));

                    var run3 = tocPara.AppendChild(new Run());
                    run3.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Separate });

                    // 添加一个空的段落作为目录占位符
                    tocPara.AppendChild(new Run(new Text("")));

                    var run4 = tocPara.AppendChild(new Run());
                    run4.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.End });

                    // 添加分页符
                    body.AppendChild(new Paragraph(
                        new Run(
                            new Break() { Type = BreakValues.Page }
                        )
                    ));

                    // 添加节属性以确保页码正确显示
                    var sectionProps = new SectionProperties(
                        new PageSize() { Width = 12240, Height = 15840 },
                        new PageMargin()
                        {
                            Top = 1440,
                            Right = 1440,
                            Bottom = 1440,
                            Left = 1440,
                            Header = 720,
                            Footer = 720
                        }
                    );
                    body.AppendChild(sectionProps);

                    // 添加各章节标题
                    var chapters = new[]
                    {
                        "1. 目的",
                        "2. 实验地点和时间",
                        "3. 实验方案",
                        "4. 结果与分析",
                        "5 结论",
                        "参考文献：",
                        "附件："
                    };

                    foreach (var chapter in chapters.Take(chapters.Length - 2)) // 跳过参考文献和附件
                    {
                        AddHeading(body, chapter, "Heading1");

                        // 如果是"结果与分析"章节，添加Excel表格和图片
                        if (chapter.Contains("结果与分析"))
                        {
                            ProcessExcelSheets(workbook, body, mainPart);
                        }

                        // 添加空白段落作为章节间隔
                        body.AppendChild(new Paragraph(
                            new ParagraphProperties(
                                new SpacingBetweenLines() { After = "800" }
                            )
                        ));
                    }

                    // 添加参考文献和附件
                    foreach (var chapter in chapters.Skip(chapters.Length - 2))
                    {
                        AddHeading(body, chapter, "Heading1");
                    }

                    // 添加样式定义
                    AddStyleDefinitions(mainPart);
                }
            }
        }
        catch (IOException ex)
        {
            throw new Exception($"文件访问错误: {ex.Message}", ex);
        }
        catch (OpenXmlPackageException ex)
        {
            throw new Exception($"Word文档创建错误: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"转换过程中出现错误: {ex.Message}", ex);
        }
    }

    public void ConvertMultipleExcelsToWord(string[] excelPaths, string wordPath)
    {
        // 检查参数
        if (excelPaths == null || excelPaths.Length == 0)
            throw new ArgumentException("Excel文件路径列表不能为空", nameof(excelPaths));
        if (string.IsNullOrEmpty(wordPath))
            throw new ArgumentException("Word文件路径不能为空", nameof(wordPath));

        // 检查所有Excel文件是否存在
        foreach (var path in excelPaths)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"找不到Excel文件: {path}", path);
        }

        try
        {
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Create(wordPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // 添加目录标题和目录
                AddChapterTitle(body, "目录", 1, true);
                AddTableOfContents(body);

                // 添加分页符
                AddPageBreak(body);

                // 添加节属性
                AddSectionProperties(body);

                // 处理每个Excel文件
                for (int fileIndex = 0; fileIndex < excelPaths.Length; fileIndex++)
                {
                    using (FileStream fs = new FileStream(excelPaths[fileIndex], FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook = new XSSFWorkbook(fs);
                        string fileName = Path.GetFileNameWithoutExtension(excelPaths[fileIndex]);

                        // 为每个Excel文件添加标题
                        AddHeading(body, $"{fileIndex + 1}. {fileName}", "Heading1");

                        // 处理工作表（包含图片）
                        ProcessExcelSheets(workbook, body, mainPart);

                        // 添加分页符（除了最后一个文件）
                        if (fileIndex < excelPaths.Length - 1)
                        {
                            AddPageBreak(body);
                        }
                    }
                }

                // 添加样式定义
                AddStyleDefinitions(mainPart);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"转换过程中出现错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 处理Excel工作表，包括表格和嵌入的图片
    /// </summary>
    private void ProcessExcelSheets(IWorkbook workbook, Body body, MainDocumentPart mainPart)
    {
        for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            if (sheet == null) continue;

            // 获取实际表格范围
            var tableRange = GetTableRange(sheet);
            if (tableRange.StartRow == tableRange.EndRow) continue; // 跳过空表格



            // 先提取所有图片信息
            var images = new List<ExcelImageInfo>();
            if (sheet is XSSFSheet xssfSheet)
            {
                images = ExtractImagesFromSheet(xssfSheet);

            }

            // 添加表格标题
            AddTableTitle(body, sheetIndex, sheet.SheetName);

            // 创建表格
            var table = CreateTable();

            // 使用实际范围填充表格数据，包括嵌入的图片
            FillTableData(table, sheet, mainPart, images,
                tableRange.StartRow, tableRange.EndRow + 1,
                tableRange.StartCol, tableRange.EndCol + 1);

            body.Append(table);

            // 在表格后添加间距
            AddSpacingAfterTable(body);
        }
    }

    /// <summary>
    /// 从Excel工作表中提取所有图片信息
    /// </summary>
    private List<ExcelImageInfo> ExtractImagesFromSheet(XSSFSheet sheet)
    {
        var images = new List<ExcelImageInfo>();
        
        try
        {
            var drawing = sheet.GetDrawingPatriarch() as XSSFDrawing;
            if (drawing == null) return images;

            // 获取所有图形对象
            var shapes = drawing.GetShapes();
            
            foreach (var shape in shapes)
            {
                if (shape is XSSFPicture picture)
                {
                    try
                    {
                        var imageInfo = ExtractImageInfo(picture);
                        if (imageInfo != null)
                        {
                            images.Add(imageInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        // 忽略单个图片提取错误
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // 忽略工作表图片提取错误
        }

        return images;
    }

    /// <summary>
    /// 提取单个图片的详细信息
    /// </summary>
    private ExcelImageInfo ExtractImageInfo(XSSFPicture picture)
    {
        try
        {
            var pictureData = picture.PictureData; // 使用属性而不是方法
            var anchor = picture.ClientAnchor; // 使用属性而不是方法
            
            return new ExcelImageInfo
            {
                ImageData = pictureData.Data,
                FileName = $"image_{Guid.NewGuid():N}.{GetImageExtension(pictureData.PictureType)}",
                Width = GetImageWidthFromAnchor(anchor),
                Height = GetImageHeightFromAnchor(anchor),
                ContentType = GetImageContentType(pictureData.PictureType),
                Row = anchor.Row1,
                Column = anchor.Col1
            };
        }
        catch (Exception ex)
        {
            // 提取图片信息失败
            return null;
        }
    }

    /// <summary>
    /// 从ClientAnchor获取图片宽度
    /// </summary>
    private double GetImageWidthFromAnchor(IClientAnchor anchor)
    {
        try
        {
            // 计算列跨度的像素宽度
            int colSpan = anchor.Col2 - anchor.Col1 + 1;
            return colSpan * 64.0; // 假设每列64像素宽度
        }
        catch
        {
            return 200.0; // 默认宽度
        }
    }

    /// <summary>
    /// 从ClientAnchor获取图片高度
    /// </summary>
    private double GetImageHeightFromAnchor(IClientAnchor anchor)
    {
        try
        {
            // 计算行跨度的像素高度
            int rowSpan = anchor.Row2 - anchor.Row1 + 1;
            return rowSpan * 20.0; // 假设每行20像素高度
        }
        catch
        {
            return 150.0; // 默认高度
        }
    }

    /// <summary>
    /// 获取图片文件扩展名 - C# 7.3兼容版本
    /// </summary>
    private string GetImageExtension(PictureType pictureType)
    {
        switch (pictureType)
        {
            case PictureType.JPEG:
                return "jpg";
            case PictureType.PNG:
                return "png";
            case PictureType.GIF:
                return "gif";
            case PictureType.BMP:
                return "bmp";
            case PictureType.TIFF:
                return "tiff";
            default:
                return "jpg";
        }
    }

    /// <summary>
    /// 将NPOI图片类型转换为OpenXml内容类型字符串 - C# 7.3兼容版本
    /// </summary>
    private string GetImageContentType(PictureType pictureType)
    {
        switch (pictureType)
        {
            case PictureType.JPEG:
                return "image/jpeg";
            case PictureType.PNG:
                return "image/png";
            case PictureType.GIF:
                return "image/gif";
            case PictureType.BMP:
                return "image/bmp";
            case PictureType.TIFF:
                return "image/tiff";
            default:
                return "image/jpeg";
        }
    }

    /// <summary>
    /// 创建OpenXml图片元素
    /// </summary>
    private Drawing CreateImageElement(string relationshipId, long widthEmu, long heightEmu, string fileName)
    {
        return new Drawing(
            new DW.Inline(
                new DW.Extent() { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = fileName
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)0U,
                                    Name = fileName
                                },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip(
                                    new A.BlipExtensionList(
                                        new A.BlipExtension()
                                        {
                                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                        })
                                )
                                {
                                    Embed = relationshipId,
                                    CompressionState = A.BlipCompressionValues.Print
                                },
                                new A.Stretch(
                                    new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = widthEmu, Cy = heightEmu }),
                                new A.PresetGeometry(
                                    new A.AdjustValueList()
                                ) { Preset = A.ShapeTypeValues.Rectangle }))
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            });
    }

    private (int StartRow, int EndRow, int StartCol, int EndCol) GetTableRange(ISheet sheet)
    {
        int startRow = -1, endRow = -1, startCol = -1, endCol = -1;

        // 遍历所有行和列来寻找带边框的单元格
        for (int row = 0; row <= sheet.LastRowNum; row++)
        {
            var excelRow = sheet.GetRow(row);
            if (excelRow == null) continue;

            for (int col = 0; col < excelRow.LastCellNum; col++)
            {
                var cell = excelRow.GetCell(col);
                if (cell == null) continue;

                // 获取单元格的样式
                var cellStyle = cell.CellStyle;
                if (cellStyle != null && HasBorder(cellStyle))
                {
                    // 更新表格范围
                    if (startRow == -1) startRow = row;
                    if (startCol == -1 || col < startCol) startCol = col;
                    if (col > endCol) endCol = col;
                    endRow = row;
                }
            }
        }

        // 如果没有找到带边框的单元格，返回默认值
        if (startRow == -1 || startCol == -1)
        {
            return (0, 0, 0, 0);
        }

        // 确保找到了完整的表格范围
        // 向外扩展搜索范围，确保包含所有关的合并单元格
        for (int i = 0; i < sheet.NumMergedRegions; i++)
        {
            var region = sheet.GetMergedRegion(i);
            if (region.FirstRow <= endRow && region.LastRow >= startRow &&
                region.FirstColumn <= endCol && region.LastColumn >= startCol)
            {
                startRow = Math.Min(startRow, region.FirstRow);
                endRow = Math.Max(endRow, region.LastRow);
                startCol = Math.Min(startCol, region.FirstColumn);
                endCol = Math.Max(endCol, region.LastColumn);
            }
        }

        return (startRow, endRow, startCol, endCol);
    }

    private bool HasBorder(ICellStyle cellStyle)
    {
        // 检查单元格是否有边框
        return cellStyle.BorderBottom != BorderStyle.None ||
               cellStyle.BorderTop != BorderStyle.None ||
               cellStyle.BorderLeft != BorderStyle.None ||
               cellStyle.BorderRight != BorderStyle.None;
    }

    private void AddTableTitle(Body body, int index, string sheetName)
    {
        var titlePara = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId() { Val = "Heading2" },  // 使用二级标题样式
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { After = "400", Before = "400" }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize() { Val = "24" },
                    new RunFonts()
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        EastAsia = "宋体",
                        ComplexScript = "Times New Roman"
                    }
                ),
                new Text($"表{index + 1} {sheetName}")
            )
        );
        body.AppendChild(titlePara);
    }

    private DocumentFormat.OpenXml.Wordprocessing.Table CreateTable()
    {
        return new DocumentFormat.OpenXml.Wordprocessing.Table(
            new TableProperties(
                new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" }
                ),
                new TableLayout() { Type = TableLayoutValues.Fixed },
                new TableLook() { Val = "04A0" }
            )
        );
    }

    private void FillTableData(DocumentFormat.OpenXml.Wordprocessing.Table table,
        ISheet sheet, MainDocumentPart mainPart, List<ExcelImageInfo> images,
        int startRow, int endRow, int startCol, int endCol)
    {
        int rowCount = endRow - startRow;
        int colCount = endCol - startCol;

        // 创建一个二维数组来跟踪合并状态
        var mergeStatus = new (bool IsMerged, bool IsFirst, bool IsVerticalMerge, int RowSpan, int ColSpan, string Content)[rowCount, colCount];

        // 首先标记所有合并单元格
        for (int i = 0; i < sheet.NumMergedRegions; i++)
        {
            var region = sheet.GetMergedRegion(i);
            if (region.FirstRow >= startRow && region.LastRow < endRow &&
                region.FirstColumn >= startCol && region.LastColumn < endCol)
            {
                // 使用DataFormatter获取格式化后的单元格值
                ICell cell = sheet.GetRow(region.FirstRow)?.GetCell(region.FirstColumn);
                string content = cell != null ? _formatter.FormatCellValue(cell) : string.Empty;

                int relativeFirstRow = region.FirstRow - startRow;
                int relativeLastRow = region.LastRow - startRow;
                int relativeFirstCol = region.FirstColumn - startCol;
                int relativeLastCol = region.LastColumn - startCol;

                bool isVerticalMerge = relativeLastRow > relativeFirstRow;

                for (int row = relativeFirstRow; row <= relativeLastRow; row++)
                {
                    for (int col = relativeFirstCol; col <= relativeLastCol; col++)
                    {
                        bool isFirst = (row == relativeFirstRow && col == relativeFirstCol);
                        int rowSpan = relativeLastRow - relativeFirstRow + 1;
                        int colSpan = relativeLastCol - relativeFirstCol + 1;
                        mergeStatus[row, col] = (true, isFirst, isVerticalMerge, rowSpan, colSpan, content);
                    }
                }
            }
        }

        // 处理每一行
        for (int i = 0; i < rowCount; i++)
        {
            TableRow tr = new TableRow(
                new TableRowProperties(
                    new TableRowHeight() { Val = 400, HeightType = HeightRuleValues.AtLeast }
                )
            );
            IRow excelRow = sheet.GetRow(i + startRow);

            // 处理每一列
            for (int j = 0; j < colCount; j++)
            {
                var currentCell = mergeStatus[i, j];

                // 在开头声明cellRow和cellCol，避免重复声明
                var cellRow = i + startRow;
                var cellCol = j + startCol;

                // 如果是水平合并但不是第一个单元格，跳过
                if (currentCell.IsMerged && !currentCell.IsFirst && !currentCell.IsVerticalMerge)
                {
                    continue;
                }

                // 创建单元格
                var tc = new TableCell();
                var tcProps = new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2000" },
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new TableCellBorders(
                        new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" },
                        new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" },
                        new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" },
                        new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" }
                    )
                );

                // 处理垂直合并
                if (currentCell.IsVerticalMerge)
                {
                    if (currentCell.IsFirst)
                    {
                        tcProps.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                        tcProps.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
                    }
                    else
                    {
                        tcProps.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                        tc.Append(tcProps);
                        tc.Append(new Paragraph(
                            new ParagraphProperties(
                                new Justification() { Val = JustificationValues.Center }
                            )
                        ));
                        tr.Append(tc);
                        continue;
                    }
                }

                // 处理水平合并
                if (currentCell.IsMerged && currentCell.ColSpan > 1)
                {
                    tcProps.Append(new GridSpan() { Val = currentCell.ColSpan });
                }

                tc.Append(tcProps);

                // 检查当前单元格是否有图片 - 完全按Excel原始位置匹配
                // 精确匹配Excel中的图片位置
                var cellImage = images.FirstOrDefault(img => img.Row == cellRow && img.Column == cellCol);
                


                if (cellImage != null)
                {
                    // 单元格包含图片，创建带图片的段落
                    try
                    {
                        var imageElement = CreateCellImageElement(mainPart, cellImage);
                        tc.Append(new Paragraph(
                            new ParagraphProperties(
                                new Justification() { Val = JustificationValues.Center },
                                new SpacingBetweenLines() { Before = "0", After = "0" }
                            ),
                            new Run(imageElement)
                        ));
                    }
                    catch (Exception ex)
                    {
                        // 图片嵌入失败，显示文本内容
                        string content = currentCell.IsMerged ? currentCell.Content : 
                            (excelRow?.GetCell(j + startCol) != null ? _formatter.FormatCellValue(excelRow.GetCell(j + startCol)) : "[图片]");
                        
                        tc.Append(new Paragraph(
                            new ParagraphProperties(
                                new Justification() { Val = JustificationValues.Center },
                                new SpacingBetweenLines() { Before = "0", After = "0" }
                            ),
                            new Run(
                                new RunProperties(
                                    new FontSize() { Val = "21" },
                                    new RunFonts()
                                    {
                                        Ascii = "Times New Roman",
                                        HighAnsi = "Times New Roman",
                                        EastAsia = "宋体",
                                        ComplexScript = "Times New Roman"
                                    }
                                ),
                                new Text(content)
                            )
                        ));
                    }
                }
                else
                {
                    // 普通单元格，添加文本内容
                    string content;
                    if (currentCell.IsMerged)
                    {
                        content = currentCell.Content;
                    }
                    else
                    {
                        ICell cell = excelRow?.GetCell(j + startCol);
                        content = cell != null ? _formatter.FormatCellValue(cell) : string.Empty;
                    }

                    tc.Append(new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center },
                            new SpacingBetweenLines() { Before = "0", After = "0" }
                        ),
                        new Run(
                            new RunProperties(
                                new FontSize() { Val = "21" },
                                new RunFonts()
                                {
                                    Ascii = "Times New Roman",
                                    HighAnsi = "Times New Roman",
                                    EastAsia = "宋体",
                                    ComplexScript = "Times New Roman"
                                }
                            ),
                            new Text(content)
                        )
                    ));
                }

                tr.Append(tc);

                // 如果是水平合并的首单元格，跳过后续的合并单元格
                if (currentCell.ColSpan > 1)
                {
                    j += currentCell.ColSpan - 1;
                }
            }
            table.Append(tr);
        }
    }

    /// <summary>
    /// 为表格单元格创建图片元素
    /// </summary>
    private Drawing CreateCellImageElement(MainDocumentPart mainPart, ExcelImageInfo imageInfo)
    {
        // 创建图片部分
        var imagePart = mainPart.AddImagePart(imageInfo.ContentType);
        
        // 写入图片数据
        using (var stream = new MemoryStream(imageInfo.ImageData))
        {
            imagePart.FeedData(stream);
        }

        // 获取图片关系ID
        var relationshipId = mainPart.GetIdOfPart(imagePart);

        // 单元格内图片使用较小尺寸（2cm x 1.5cm）
        long widthEmu = 720000L;   // 2cm
        long heightEmu = 540000L;  // 1.5cm

        // 创建图片元素
        return CreateImageElement(relationshipId, widthEmu, heightEmu, imageInfo.FileName);
    }

    private void AddSpacingAfterTable(Body body)
    {
        body.AppendChild(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines() { After = "800" }
            )
        ));
    }

    private void AddChapterTitle(Body body, string title, int fontSize, bool isBold = false)
    {
        var properties = new ParagraphProperties(
            new SpacingBetweenLines() { After = "400", Before = "400" }
        );

        var runProperties = new RunProperties(
            new FontSize() { Val = (fontSize * 12).ToString() }
        );

        if (isBold)
        {
            runProperties.AppendChild(new Bold());
        }

        var paragraph = new Paragraph(
            properties,
            new Run(
                runProperties,
                new Text(title)
            )
        );

        body.AppendChild(paragraph);
    }

    private void AddTocEntry(Body body, string entry)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines() { After = "200" }
            ),
            new Run(
                new RunProperties(
                    new FontSize() { Val = "24" }
                ),
                new Text(entry)
            )
        );

        body.AppendChild(paragraph);
    }

    private void AddHeading(Body body, string title, string styleId)
    {
        var para = body.AppendChild(new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId() { Val = styleId },
                new SpacingBetweenLines() { After = "400", Before = "400" },
                new OutlineLevel() { Val = 0 }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize() { Val = "28" }
                ),
                new Text(title)
            )
        ));
    }

    private void AddStyleDefinitions(MainDocumentPart mainPart)
    {
        // 添加样式部分如果不存在
        StyleDefinitionsPart styleDefinitionsPart;
        if (mainPart.StyleDefinitionsPart == null)
        {
            styleDefinitionsPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var initialStyles = new Styles();
            initialStyles.Save(styleDefinitionsPart);
        }
        else
        {
            styleDefinitionsPart = mainPart.StyleDefinitionsPart;
        }

        // 获取现有样式或创建新的样式集合
        var currentStyles = styleDefinitionsPart.Styles ?? new Styles();

        // 添加默认段落样式
        var defaultStyle = new Style()
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
        defaultStyle.Append(
            new Name() { Val = "Normal" },
            new RunProperties(
                new RunFonts()
                {
                    Ascii = "Times New Roman",      // 英文字体
                    HighAnsi = "Times New Roman",   // 英文字体
                    EastAsia = "宋体",              // 中文字体
                    ComplexScript = "Times New Roman"
                }
            )
        );

        // 创建标题1样式
        var heading1Style = new Style()
        {
            Type = StyleValues.Paragraph,
            StyleId = "Heading1",
            CustomStyle = true,
            Default = false
        };

        heading1Style.Append(
            new Name() { Val = "heading 1" },
            new BasedOn() { Val = "Normal" },
            new NextParagraphStyle() { Val = "Normal" },
            new ParagraphProperties(
                new SpacingBetweenLines() { After = "400", Before = "400" },
                new OutlineLevel() { Val = 0 }
            ),
            new RunProperties(
                new Bold(),
                new FontSize() { Val = "28" },
                new RunFonts()
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    EastAsia = "宋体",
                    ComplexScript = "Times New Roman"
                }
            )
        );

        // 添加二级标题样式
        var heading2Style = new Style()
        {
            Type = StyleValues.Paragraph,
            StyleId = "Heading2",
            CustomStyle = true,
            Default = false
        };

        heading2Style.Append(
            new Name() { Val = "heading 2" },
            new BasedOn() { Val = "Normal" },
            new NextParagraphStyle() { Val = "Normal" },
            new ParagraphProperties(
                new SpacingBetweenLines() { After = "400", Before = "400" },
                new OutlineLevel() { Val = 1 }  // 设置为二级标题级别
            ),
            new RunProperties(
                new Bold(),
                new FontSize() { Val = "24" },
                new RunFonts()
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    EastAsia = "宋体",
                    ComplexScript = "Times New Roman"
                }
            )
        );

        // 更新或添加样式
        var existingDefaultStyle = currentStyles.Elements<Style>().FirstOrDefault(s => s.StyleId == "Normal");
        if (existingDefaultStyle != null)
        {
            existingDefaultStyle.Remove();
        }
        currentStyles.Append(defaultStyle);

        var existingHeading1Style = currentStyles.Elements<Style>().FirstOrDefault(s => s.StyleId == "Heading1");
        if (existingHeading1Style != null)
        {
            existingHeading1Style.Remove();
        }
        currentStyles.Append(heading1Style);

        // 更新或添加二级标题样式
        var existingHeading2Style = currentStyles.Elements<Style>().FirstOrDefault(s => s.StyleId == "Heading2");
        if (existingHeading2Style != null)
        {
            existingHeading2Style.Remove();
        }
        currentStyles.Append(heading2Style);

        // 保存样式
        currentStyles.Save(styleDefinitionsPart);
    }

    private void AddTableOfContents(Body body)
    {
        var tocPara = body.AppendChild(new Paragraph());
        var run = tocPara.AppendChild(new Run());
        run.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Begin });

        var run2 = tocPara.AppendChild(new Run());
        run2.AppendChild(new Text(" TOC \\h \\z "));

        var run3 = tocPara.AppendChild(new Run());
        run3.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Separate });

        tocPara.AppendChild(new Run(new Text("")));

        var run4 = tocPara.AppendChild(new Run());
        run4.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.End });
    }

    private void AddPageBreak(Body body)
    {
        body.AppendChild(new Paragraph(
            new Run(
                new Break() { Type = BreakValues.Page }
            )
        ));
    }

    private void AddSectionProperties(Body body)
    {
        var sectionProps = new SectionProperties(
            new PageSize() { Width = 12240, Height = 15840 },
            new PageMargin()
            {
                Top = 1440,
                Right = 1440,
                Bottom = 1440,
                Left = 1440,
                Header = 720,
                Footer = 720
            }
        );
        body.AppendChild(sectionProps);
    }
}

/// <summary>
/// Excel图片信息类
/// </summary>
public class ExcelImageInfo
{
    /// <summary>
    /// 图片二进制数据
    /// </summary>
    public byte[] ImageData { get; set; }

    /// <summary>
    /// 文件名
    /// </summary>
    public string FileName { get; set; }

    /// <summary>
    /// 宽度（像素）
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// 高度（像素）
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// OpenXml内容类型
    /// </summary>
    public string ContentType { get; set; }

    /// <summary>
    /// 图片在Excel中的行位置
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    /// 图片在Excel中的列位置
    /// </summary>
    public int Column { get; set; }
}