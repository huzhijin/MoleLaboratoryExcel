using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
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

                    // 添加节属性，使用横向布局以容纳宽表格
                    var sectionProps = new SectionProperties(
                        new PageSize() { Width = 15840, Height = 12240, Orient = PageOrientationValues.Landscape },
                        new PageMargin()
                        {
                            Top = 720,
                            Right = 720,
                            Bottom = 720,
                            Left = 720,
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
                var rawImages = ExtractImagesFromSheet(xssfSheet);
                // 进一步过滤和去重，处理WPS可能产生的问题
                images = FilterAndDeduplicateImages(rawImages);
            }

            // 添加表格标题
            AddTableTitle(body, sheetIndex, sheet.SheetName);

            // 创建表格
            var table = CreateTable();

            // 显示图片分布概览
            System.Diagnostics.Debug.WriteLine($"📊 表格处理概览:");
            System.Diagnostics.Debug.WriteLine($"   表格范围: 行{tableRange.StartRow}-{tableRange.EndRow}, 列{tableRange.StartCol}-{tableRange.EndCol}");
            System.Diagnostics.Debug.WriteLine($"   图片总数: {images.Count}张");
            foreach (var img in images)
            {
                string cropInfo = img.HasCropping ? $"L{img.CropLeft:F0}%T{img.CropTop:F0}%R{img.CropRight:F0}%B{img.CropBottom:F0}%" : "无裁剪";
                System.Diagnostics.Debug.WriteLine($"   📍 图片{img.ImageIndex}: 位置({img.StartRow},{img.StartCol}), 大小{img.ImageData.Length/1024}KB, {cropInfo}");
            }
            
            // 特别检查K列图片
            var kColumnImages = images.Where(img => img.StartCol == 10).ToList();
            if (kColumnImages.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"🔍 K列(第10列)图片: {kColumnImages.Count}张");
                foreach (var img in kColumnImages)
                {
                    System.Diagnostics.Debug.WriteLine($"   📊 ROC图{img.ImageIndex}: 行{img.StartRow}, 大小{img.ImageData.Length/1024}KB");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"❌ K列(第10列)没有找到图片！");
                // 显示所有图片的列分布
                var columnDistribution = images.GroupBy(img => img.StartCol).OrderBy(g => g.Key);
                System.Diagnostics.Debug.WriteLine($"🗺️ 图片列分布:");
                foreach (var group in columnDistribution)
                {
                    System.Diagnostics.Debug.WriteLine($"   列{group.Key}: {group.Count()}张图片");
                }
            }
            
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
            
            int imageIndex = 0;
            foreach (var shape in shapes)
            {
                if (shape is XSSFPicture picture)
                {
                    try
                    {
                        var imageInfo = ExtractImageInfo(picture, imageIndex);
                        if (imageInfo != null)
                        {
                            // 移除所有过滤条件：Excel里有的图片都要传过来
                            // 用户要求：不用过滤，只要Excel里有的就传过来
                            if (imageInfo.ImageData.Length > 0) // 只要有数据就保留
                            {
                                images.Add(imageInfo);
                                System.Diagnostics.Debug.WriteLine($"提取图片 {imageIndex}: 位置({imageInfo.StartRow}, {imageInfo.StartCol}) 到 ({imageInfo.EndRow}, {imageInfo.EndCol}), 大小: {imageInfo.ImageData.Length} bytes, 尺寸: {imageInfo.Width}x{imageInfo.Height}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"跳过图片 {imageIndex}: 无数据, 大小: {imageInfo.ImageData.Length} bytes");
                            }
                        }
                        imageIndex++;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"提取图片 {imageIndex} 失败: {ex.Message}");
                        imageIndex++;
                    }
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"工作表 '{sheet.SheetName}' 总共找到 {images.Count} 张有效图片");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"提取工作表图片失败: {ex.Message}");
        }

        return images;
    }

    /// <summary>
    /// 过滤和去重图片，处理WPS兼容性问题
    /// </summary>
    private List<ExcelImageInfo> FilterAndDeduplicateImages(List<ExcelImageInfo> rawImages)
    {
        var filteredImages = new List<ExcelImageInfo>();
        
        if (rawImages == null || rawImages.Count == 0)
            return filteredImages;
        
        // 按位置分组，处理同一位置的多张图片
        var groupedByPosition = rawImages.GroupBy(img => new { img.StartRow, img.StartCol });
        
        foreach (var group in groupedByPosition)
        {
            var imagesAtPosition = group.ToList();
            
            if (imagesAtPosition.Count == 1)
            {
                // 只有一张图片，直接添加
                filteredImages.Add(imagesAtPosition[0]);
                System.Diagnostics.Debug.WriteLine($"位置({group.Key.StartRow},{group.Key.StartCol})：保留唯一图片{imagesAtPosition[0].ImageIndex}");
            }
            else
            {
                // 多张图片在同一位置，需要智能选择
                var selectedImage = SelectBestImageFromGroup(imagesAtPosition);
                if (selectedImage != null)
                {
                    filteredImages.Add(selectedImage);
                    System.Diagnostics.Debug.WriteLine($"位置({group.Key.StartRow},{group.Key.StartCol})：从{imagesAtPosition.Count}张图片中选择图片{selectedImage.ImageIndex}");
                }
            }
        }
        
        System.Diagnostics.Debug.WriteLine($"图片过滤完成：原始{rawImages.Count}张，过滤后{filteredImages.Count}张");
        return filteredImages;
    }

    /// <summary>
    /// 从同一位置的多张图片中选择最佳的一张
    /// </summary>
    private ExcelImageInfo SelectBestImageFromGroup(List<ExcelImageInfo> images)
    {
        if (images == null || images.Count == 0)
            return null;
            
        // 移除过滤条件：保留所有图片
        // 用户要求：Excel里什么样就什么样，不过滤
        var validImages = images.ToList(); // 保留所有图片
        
        // 选择策略：
        // 1. 优先选择数据量最大的图片（通常质量更好）
        // 2. 其次选择尺寸最大的图片
        var selectedImage = validImages
            .OrderByDescending(img => img.ImageData.Length)
            .ThenByDescending(img => img.Width * img.Height)
            .First();
            
        System.Diagnostics.Debug.WriteLine($"从{images.Count}张候选图片中选择：图片{selectedImage.ImageIndex}，数据大小{selectedImage.ImageData.Length}字节，尺寸{selectedImage.Width}x{selectedImage.Height}");
        
        return selectedImage;
    }

    /// <summary>
    /// 提取单个图片的详细信息，包括裁剪信息
    /// </summary>
    private ExcelImageInfo ExtractImageInfo(XSSFPicture picture, int imageIndex)
    {
        try
        {
            var pictureData = picture.PictureData;
            var anchor = picture.ClientAnchor;
            
            // 验证图片数据是否有效
            if (pictureData == null || pictureData.Data == null || pictureData.Data.Length == 0)
            {
                return null;
            }
            
            // 验证锚点是否有效
            if (anchor == null)
            {
                return null;
            }
            
            // 提取裁剪信息
            var cropInfo = ExtractCroppingInfo(picture);
            
            System.Diagnostics.Debug.WriteLine($"图片{imageIndex} - 原始位置: ({anchor.Row1},{anchor.Col1}) 到 ({anchor.Row2},{anchor.Col2})");
            if (cropInfo.HasCropping)
            {
                System.Diagnostics.Debug.WriteLine($"图片{imageIndex} - 裁剪信息: Left={cropInfo.CropLeft}, Top={cropInfo.CropTop}, Right={cropInfo.CropRight}, Bottom={cropInfo.CropBottom}");
            }
            
            return new ExcelImageInfo
            {
                ImageData = pictureData.Data,
                FileName = $"image_{imageIndex}_{Guid.NewGuid():N}.{GetImageExtension(pictureData.PictureType)}",
                Width = GetImageWidthFromAnchor(anchor),
                Height = GetImageHeightFromAnchor(anchor),
                ContentType = GetImageContentType(pictureData.PictureType),
                Row = anchor.Row1,
                Column = anchor.Col1,
                // 添加图片范围信息，用于更准确的位置匹配
                StartRow = anchor.Row1,
                EndRow = anchor.Row2,
                StartCol = anchor.Col1,
                EndCol = anchor.Col2,
                ImageIndex = imageIndex,
                // 添加裁剪信息
                CropLeft = cropInfo.CropLeft,
                CropTop = cropInfo.CropTop,
                CropRight = cropInfo.CropRight,
                CropBottom = cropInfo.CropBottom,
                HasCropping = cropInfo.HasCropping
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"提取图片{imageIndex}信息失败: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 提取图片的裁剪信息 - 启用OpenXML标准裁剪，禁用自动检测
    /// </summary>
    private (bool HasCropping, double CropLeft, double CropTop, double CropRight, double CropBottom) ExtractCroppingInfo(XSSFPicture picture)
    {
        try
        {
            // 只使用Excel中真正的裁剪信息（OpenXML标准）
            var ctPicture = picture.GetCTPicture();
            if (ctPicture?.blipFill?.srcRect != null)
            {
                var srcRect = ctPicture.blipFill.srcRect;
                
                // 转换为百分比值（OpenXML中以千分比存储）
                double cropLeft = srcRect.l / 1000.0;   
                double cropTop = srcRect.t / 1000.0;    
                double cropRight = srcRect.r / 1000.0;  
                double cropBottom = srcRect.b / 1000.0; 
                
                bool hasCropping = cropLeft > 0 || cropTop > 0 || cropRight > 0 || cropBottom > 0;
                
                System.Diagnostics.Debug.WriteLine($"检测到Excel标准裁剪: L={cropLeft:F1}% T={cropTop:F1}% R={cropRight:F1}% B={cropBottom:F1}%");
                
                return (hasCropping, cropLeft, cropTop, cropRight, cropBottom);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("未检测到Excel裁剪信息");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"提取裁剪信息失败: {ex.Message}");
        }
        
        return (false, 0, 0, 0, 0);
    }

    /// <summary>
    /// 检查图片是否被单元格边界裁剪
    /// </summary>
    private bool CheckCellClipping(XSSFPicture picture, NPOI.XSSF.UserModel.XSSFClientAnchor anchor)
    {
        try
        {
            // 获取图片的实际像素尺寸和在Excel中的显示尺寸
            var pictureData = picture.PictureData;
            if (pictureData?.Data == null) return false;
            
            using (var stream = new MemoryStream(pictureData.Data))
            using (var img = Image.FromStream(stream))
            {
                int actualWidth = img.Width;
                int actualHeight = img.Height;
                
                // 计算Excel中的显示区域大小（像素）
                double displayWidth = GetImageWidthFromAnchor(anchor);
                double displayHeight = GetImageHeightFromAnchor(anchor);
                
                // 如果实际图片明显大于显示区域，可能存在裁剪
                bool widthClipped = actualWidth > displayWidth * 2.0; // 允许100%的容差，更保守
                bool heightClipped = actualHeight > displayHeight * 2.0;
                
                System.Diagnostics.Debug.WriteLine($"检查单元格裁剪 - 实际尺寸: {actualWidth}x{actualHeight}, 显示尺寸: {displayWidth:F0}x{displayHeight:F0}, 裁剪: {widthClipped || heightClipped}");
                
                return widthClipped || heightClipped;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"检查单元格裁剪失败: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 计算单元格裁剪比例
    /// </summary>
    private (double Left, double Top, double Right, double Bottom) CalculateCellCropping(XSSFPicture picture, NPOI.XSSF.UserModel.XSSFClientAnchor anchor)
    {
        try
        {
            var pictureData = picture.PictureData;
            if (pictureData?.Data == null) return (0, 0, 0, 0);
            
            using (var stream = new MemoryStream(pictureData.Data))
            using (var img = Image.FromStream(stream))
            {
                int actualWidth = img.Width;
                int actualHeight = img.Height;
                
                // 计算Excel中的显示区域大小
                double displayWidth = GetImageWidthFromAnchor(anchor);
                double displayHeight = GetImageHeightFromAnchor(anchor);
                
                // 计算裁剪比例
                double cropLeft = 0;
                double cropTop = 0;
                double cropRight = 0;
                double cropBottom = 0;
                
                // 如果图片宽度被裁剪（更保守的判断）
                if (actualWidth > displayWidth * 1.8)  // 只有明显超出时才裁剪
                {
                    double widthRatio = displayWidth / actualWidth;
                    cropRight = Math.Min(50.0, (1.0 - widthRatio) * 100.0); // 最多裁剪50%
                }
                
                // 如果图片高度被裁剪（更保守的判断）
                if (actualHeight > displayHeight * 1.8)  // 只有明显超出时才裁剪
                {
                    double heightRatio = displayHeight / actualHeight;
                    cropBottom = Math.Min(50.0, (1.0 - heightRatio) * 100.0); // 最多裁剪50%
                }
                
                System.Diagnostics.Debug.WriteLine($"计算单元格裁剪 - 实际尺寸: {actualWidth}x{actualHeight}, 显示尺寸: {displayWidth:F0}x{displayHeight:F0}");
                System.Diagnostics.Debug.WriteLine($"裁剪比例: 左{cropLeft}% 上{cropTop}% 右{cropRight:F1}% 下{cropBottom:F1}%");
                
                return (cropLeft, cropTop, cropRight, cropBottom);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"计算单元格裁剪失败: {ex.Message}");
            return (0, 0, 0, 0);
        }
    }

    /// <summary>
    /// 从ClientAnchor获取图片宽度 - 改进版本，兼容WPS
    /// </summary>
    private double GetImageWidthFromAnchor(IClientAnchor anchor)
    {
        try
        {
            if (anchor is NPOI.XSSF.UserModel.XSSFClientAnchor xssfAnchor)
            {
                // 使用更精确的坐标计算
                double width = 0;
                
                // 如果有精确的像素坐标，使用像素坐标
                if (xssfAnchor.Dx2 > 0 || xssfAnchor.Dx1 > 0)
                {
                    width = Math.Abs(xssfAnchor.Dx2 - xssfAnchor.Dx1) / 9525.0; // EMU to pixels
                }
                
                // 如果像素坐标无效，回退到列计算
                if (width <= 0)
                {
                    int colSpan = Math.Max(1, anchor.Col2 - anchor.Col1);
                    width = colSpan * 64.0; // 默认列宽
                }
                
                return Math.Max(width, 50.0); // 最小宽度50像素
            }
            
            // 非XSSF格式的回退处理
            int defaultColSpan = Math.Max(1, anchor.Col2 - anchor.Col1);
            return defaultColSpan * 64.0;
        }
        catch
        {
            return 200.0; // 默认宽度
        }
    }

    /// <summary>
    /// 从ClientAnchor获取图片高度 - 改进版本，兼容WPS
    /// </summary>
    private double GetImageHeightFromAnchor(IClientAnchor anchor)
    {
        try
        {
            if (anchor is NPOI.XSSF.UserModel.XSSFClientAnchor xssfAnchor)
            {
                // 使用更精确的坐标计算
                double height = 0;
                
                // 如果有精确的像素坐标，使用像素坐标
                if (xssfAnchor.Dy2 > 0 || xssfAnchor.Dy1 > 0)
                {
                    height = Math.Abs(xssfAnchor.Dy2 - xssfAnchor.Dy1) / 9525.0; // EMU to pixels
                }
                
                // 如果像素坐标无效，回退到行计算
                if (height <= 0)
                {
                    int rowSpan = Math.Max(1, anchor.Row2 - anchor.Row1);
                    height = rowSpan * 20.0; // 默认行高
                }
                
                return Math.Max(height, 30.0); // 最小高度30像素
            }
            
            // 非XSSF格式的回退处理
            int defaultRowSpan = Math.Max(1, anchor.Row2 - anchor.Row1);
            return defaultRowSpan * 20.0;
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

        // 如果没有找到带边框的单元格，使用默认范围但确保包含图片
        if (startRow == -1 || startCol == -1)
        {
            System.Diagnostics.Debug.WriteLine($"⚠️ 未找到带边框单元格，使用默认范围");
            startRow = 0;
            endRow = Math.Max(5, sheet.LastRowNum); // 至少5行
            startCol = 0;
            endCol = 10; // 至少包含到K列
        }

        // 确保找到了完整的表格范围
        // 向外扩展搜索范围，确保包含所有相关的合并单元格
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

        // 重要：检查是否有图片，如果有，扩展表格范围以包含所有图片
        if (sheet is XSSFSheet xssfSheet)
        {
            var allImages = ExtractImagesFromSheet(xssfSheet);
            System.Diagnostics.Debug.WriteLine($"📊 工作表 '{sheet.SheetName}' 发现 {allImages.Count} 张图片");
            
            if (allImages.Count > 0)
            {
                foreach (var image in allImages)
                {
                    System.Diagnostics.Debug.WriteLine($"🖼️ 图片{image.ImageIndex}位置: ({image.StartRow},{image.StartCol}) 到 ({image.EndRow},{image.EndCol})");
                    
                    // 无条件扩展表格范围以包含所有图片
                    int oldStartCol = startCol, oldEndCol = endCol, oldStartRow = startRow, oldEndRow = endRow;
                    
                    startCol = Math.Min(startCol, image.StartCol);
                    endCol = Math.Max(endCol, image.EndCol);
                    startRow = Math.Min(startRow, image.StartRow);
                    endRow = Math.Max(endRow, image.EndRow);
                    
                    if (oldStartCol != startCol || oldEndCol != endCol || oldStartRow != startRow || oldEndRow != endRow)
                    {
                        System.Diagnostics.Debug.WriteLine($"📈 范围扩展: ({oldStartRow},{oldStartCol})-({oldEndRow},{oldEndCol}) → ({startRow},{startCol})-({endRow},{endCol})");
                    }
                }
                
                // 强制确保包含K列（第10列）如果有图片在那里
                var maxImageCol = allImages.Max(img => img.StartCol);
                if (maxImageCol >= 10)
                {
                    endCol = Math.Max(endCol, maxImageCol);
                    System.Diagnostics.Debug.WriteLine($"🎯 强制包含到第{maxImageCol}列，确保ROC曲线图显示");
                }
            }
        }

        System.Diagnostics.Debug.WriteLine($"工作表 '{sheet.SheetName}' 表格范围: ({startRow},{startCol}) 到 ({endRow},{endCol})");
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
                new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto }, // 自动宽度，根据列宽计算
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Space = 0, Color = "000000" }
                ),
                new TableLayout() { Type = TableLayoutValues.Fixed }, // 保持固定布局以确保列宽精确控制
                new TableLook() { Val = "04A0" },
                new TableJustification() { Val = TableRowAlignmentValues.Left } // 表格左对齐，避免居中导致截断
            )
        );
    }

    private void FillTableData(DocumentFormat.OpenXml.Wordprocessing.Table table,
        ISheet sheet, MainDocumentPart mainPart, List<ExcelImageInfo> images,
        int startRow, int endRow, int startCol, int endCol)
    {
        int rowCount = endRow - startRow;
        int colCount = endCol - startCol;
        
        System.Diagnostics.Debug.WriteLine($"🏗️ 开始填充表格数据:");
        System.Diagnostics.Debug.WriteLine($"   行范围: {startRow}-{endRow-1} (共{rowCount}行)");
        System.Diagnostics.Debug.WriteLine($"   列范围: {startCol}-{endCol-1} (共{colCount}列)");
        System.Diagnostics.Debug.WriteLine($"   期望包含图片: {images.Count}张");
        
        // 动态读取Excel中的列宽信息
        var columnWidths = new double[colCount];
        for (int col = 0; col < colCount; col++)
        {
            int excelCol = col + startCol;
            // NPOI中列宽以256为单位，需要转换为实际宽度
            double excelWidth = sheet.GetColumnWidth(excelCol) / 256.0; // 字符宽度
            double widthCm = excelWidth * 0.18; // 近似转换为厘米 (1字符≈0.18cm)
            columnWidths[col] = Math.Max(widthCm, 1.0); // 最小1cm
            System.Diagnostics.Debug.WriteLine($"📏 列{excelCol}Excel宽度: {excelWidth:F1}字符 → {widthCm:F1}cm");
        }

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
            var actualRowIndex = i + startRow;
            System.Diagnostics.Debug.WriteLine($"🔄 处理第{i}行 (Excel行{actualRowIndex})");
            
            // 动态读取Excel中的行高
            IRow excelRow = sheet.GetRow(actualRowIndex);
            double excelRowHeight = 0.4; // 默认行高(cm)
            if (excelRow != null)
            {
                // NPOI中行高以点(Point)为单位，1点=1/72英寸=0.0353cm
                double heightInPoints = excelRow.HeightInPoints > 0 ? excelRow.HeightInPoints : 12.75; // 默认12.75点
                excelRowHeight = heightInPoints * 0.0353; // 转换为厘米
            }
            
            // 检查这一行是否有图片
            var rowImages = images.Where(img => img.StartRow == actualRowIndex).ToList();
            if (rowImages.Count > 0)
            {
                // 对于WPS嵌入单元格的图片，保持Excel原始行高
                // 图片会自动适配行高，而不是行高适配图片
                System.Diagnostics.Debug.WriteLine($"📐 图片行保持Excel高度: {excelRowHeight:F1}cm (WPS嵌入模式)");
            }
            
            // 转换为DXA单位 (1cm ≈ 567 DXA)
            var rowHeightDxa = (uint)(excelRowHeight * 567);
            
            System.Diagnostics.Debug.WriteLine($"📐 行{actualRowIndex}Excel高度: {excelRowHeight:F1}cm → {rowHeightDxa}DXA");
            
            TableRow tr = new TableRow(
                new TableRowProperties(
                    new TableRowHeight() { Val = rowHeightDxa, HeightType = HeightRuleValues.AtLeast }
                )
            );
            
            if (rowImages.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"📸 第{actualRowIndex}行有{rowImages.Count}张图片: {string.Join(",", rowImages.Select(img => $"图片{img.ImageIndex}"))}");
            }

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

                // 检查当前单元格是否有图片 - 精确匹配位置
                var candidateImages = images.Where(img => 
                    img.StartRow == cellRow && img.StartCol == cellCol).ToList();

                // 使用Excel中的实际列宽
                double excelWidthCm = columnWidths[j];
                var isImageCell = candidateImages.Count > 0;
                
                // 对于WPS嵌入单元格的图片，保持Excel原始单元格尺寸
                // 图片会自动适配单元格，而不是单元格适配图片
                if (isImageCell)
                {
                    System.Diagnostics.Debug.WriteLine($"🖼️ 图片单元格保持Excel尺寸: {excelWidthCm:F1}cm (WPS嵌入模式)");
                }
                
                // 转换为DXA单位 (1cm ≈ 567 DXA)
                var cellWidthDxa = ((int)(excelWidthCm * 567)).ToString();
                
                System.Diagnostics.Debug.WriteLine($"📏 单元格({cellRow},{cellCol}) Excel宽度: {excelWidthCm:F1}cm → {cellWidthDxa}DXA {(isImageCell ? "(图片单元格)" : "")}");
                
                
                // 创建单元格
                var tc = new TableCell();
                var tcProps = new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidthDxa },
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new TableCellBorders(
                        new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" },
                        new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" },
                        new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" },
                        new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12, Color = "000000" }
                    )
                );
                
                // 如果是图片单元格，添加特殊属性确保不会压缩内容
                if (isImageCell)
                {
                    // 添加单元格适配内容的属性
                    tcProps.Append(new NoWrap() { Val = OnOffOnlyValues.Off });
                    tcProps.Append(new TableCellFitText() { Val = OnOffOnlyValues.Off });
                }

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
                
                ExcelImageInfo cellImage = null;
                if (candidateImages.Count > 0)
                {
                    // 选择第一个匹配的图片（保持原有顺序）
                    cellImage = candidateImages.First();
                    
                    System.Diagnostics.Debug.WriteLine($"单元格({cellRow},{cellCol})找到{candidateImages.Count}张图片，选择图片{cellImage.ImageIndex}，大小{cellImage.ImageData.Length}字节");
                }
                else
                {
                    // 调试：显示附近的图片位置和当前处理的单元格范围
                    if (cellCol >= 10) // 特别关注K列(第10列)
                    {
                        System.Diagnostics.Debug.WriteLine($"处理K列单元格({cellRow},{cellCol})，无图片匹配");
                        System.Diagnostics.Debug.WriteLine($"当前表格范围: 行{startRow}-{endRow-1}, 列{startCol}-{endCol-1}");
                        
                        var allImagesInColumn = images.Where(img => img.StartCol == cellCol).ToList();
                        if (allImagesInColumn.Count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"第{cellCol}列有{allImagesInColumn.Count}张图片:");
                            foreach (var img in allImagesInColumn)
                            {
                                System.Diagnostics.Debug.WriteLine($"  图片{img.ImageIndex}行{img.StartRow}");
                            }
                        }
                    }
                    
                    var nearbyImages = images.Where(img => 
                        Math.Abs(img.StartRow - cellRow) <= 1 && Math.Abs(img.StartCol - cellCol) <= 1).ToList();
                    if (nearbyImages.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"单元格({cellRow},{cellCol})附近有{nearbyImages.Count}张图片:");
                        foreach (var nearby in nearbyImages)
                        {
                            System.Diagnostics.Debug.WriteLine($"  图片{nearby.ImageIndex}位置({nearby.StartRow},{nearby.StartCol})");
                        }
                    }
                }
                


                if (cellImage != null)
                {
                    // 单元格包含图片，创建带图片的段落
                    System.Diagnostics.Debug.WriteLine($"🖼️ 开始处理单元格({cellRow},{cellCol})的图片{cellImage.ImageIndex}");
                    try
                    {
                        // 对于WPS嵌入单元格的图片，传入单元格尺寸让图片适配
                        var imageElement = CreateCellImageElementWithCellSize(mainPart, cellImage, excelWidthCm, excelRowHeight);
                        
                        // 为图片创建专门的段落，确保不会被压缩
                        var imageParagraph = new Paragraph(
                            new ParagraphProperties(
                                new Justification() { Val = JustificationValues.Center },
                                new SpacingBetweenLines() { Before = "0", After = "0" },
                                // 确保段落不会压缩内容
                                new ContextualSpacing() { Val = false }
                            ),
                            new Run(
                                new RunProperties(
                                    // 确保运行不会自动调整大小
                                    new NoProof()
                                ),
                                imageElement
                            )
                        );
                        
                        tc.Append(imageParagraph);
                        System.Diagnostics.Debug.WriteLine($"✅ 图片{cellImage.ImageIndex}成功添加到单元格({cellRow},{cellCol})");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"❌ 图片{cellImage.ImageIndex}嵌入失败: {ex.Message}");
                        
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
            System.Diagnostics.Debug.WriteLine($"✅ 第{i}行 (Excel行{actualRowIndex}) 处理完成，包含{tr.Elements<TableCell>().Count()}个单元格");
        }
        
        System.Diagnostics.Debug.WriteLine($"🏁 表格填充完成：总共创建了{rowCount}行，期望图片{images.Count}张");
    }

    /// <summary>
    /// 为WPS嵌入单元格创建图片元素，图片适配单元格尺寸
    /// </summary>
    private Drawing CreateCellImageElementWithCellSize(MainDocumentPart mainPart, ExcelImageInfo imageInfo, double cellWidthCm, double cellHeightCm)
    {
        System.Diagnostics.Debug.WriteLine($"🔧 CreateCellImageElementWithCellSize 开始处理图片{imageInfo.ImageIndex}，单元格尺寸{cellWidthCm:F1}x{cellHeightCm:F1}cm");
        
        // 处理图片数据 - 如果有裁剪信息，先裁剪图片
        byte[] finalImageData = imageInfo.ImageData;
        
        if (imageInfo.HasCropping)
        {
            System.Diagnostics.Debug.WriteLine($"✂️ 检测到图片裁剪信息: 左{imageInfo.CropLeft:F1}% 上{imageInfo.CropTop:F1}% 右{imageInfo.CropRight:F1}% 下{imageInfo.CropBottom:F1}%");
            
            try
            {
                // 尝试裁剪图片数据
                finalImageData = CropImageData(imageInfo.ImageData, imageInfo.CropLeft, imageInfo.CropTop, imageInfo.CropRight, imageInfo.CropBottom);
                System.Diagnostics.Debug.WriteLine($"✂️ 图片裁剪成功，数据大小: {finalImageData.Length} bytes");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ 图片裁剪失败，使用原图: {ex.Message}");
                finalImageData = imageInfo.ImageData;
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"📷 图片无裁剪信息，使用原图");
        }
        
        // 创建图片部分，根据图片类型选择合适的ImagePart类型
        ImagePart imagePart;
        
        try
        {
            imagePart = mainPart.AddImagePart(imageInfo.ContentType);
        }
        catch
        {
            // 如果内容类型有问题，尝试使用默认的jpeg类型
            imagePart = mainPart.AddImagePart("image/jpeg");
        }
        
        // 写入处理后的图片数据
        using (var stream = new MemoryStream(finalImageData))
        {
            imagePart.FeedData(stream);
        }

        // 获取图片关系ID
        var relationshipId = mainPart.GetIdOfPart(imagePart);

        // WPS嵌入模式：图片直接使用单元格尺寸，不考虑图片原始尺寸
        double finalWidthCm = cellWidthCm - 0.2;  // 减去一点边距
        double finalHeightCm = cellHeightCm - 0.2; // 减去一点边距
        
        // 确保最小尺寸
        finalWidthCm = Math.Max(finalWidthCm, 0.5);
        finalHeightCm = Math.Max(finalHeightCm, 0.5);
        
        // 转换为EMU（OpenXML标准单位）
        long widthEmu = (long)(finalWidthCm * 360000);   // 1cm = 360000 EMU
        long heightEmu = (long)(finalHeightCm * 360000);
        
        System.Diagnostics.Debug.WriteLine($"WPS嵌入模式图片尺寸:");
        System.Diagnostics.Debug.WriteLine($"  单元格尺寸: {cellWidthCm:F2}x{cellHeightCm:F2}cm");
        System.Diagnostics.Debug.WriteLine($"  图片最终尺寸: {finalWidthCm:F2}x{finalHeightCm:F2}cm");
        System.Diagnostics.Debug.WriteLine($"  EMU尺寸: {widthEmu}x{heightEmu}");

        // 创建图片元素
        var drawing = CreateImageElement(relationshipId, widthEmu, heightEmu, imageInfo.FileName);
        System.Diagnostics.Debug.WriteLine($"✅ CreateCellImageElementWithCellSize 完成处理图片{imageInfo.ImageIndex}");
        return drawing;
    }

    /// <summary>
    /// 为表格单元格创建图片元素，支持裁剪
    /// </summary>
    private Drawing CreateCellImageElement(MainDocumentPart mainPart, ExcelImageInfo imageInfo)
    {
        System.Diagnostics.Debug.WriteLine($"🔧 CreateCellImageElement 开始处理图片{imageInfo.ImageIndex}，位置({imageInfo.StartRow},{imageInfo.StartCol})，数据大小{imageInfo.ImageData.Length}bytes");
        
        // 处理图片数据 - 如果有裁剪信息，先裁剪图片
        byte[] finalImageData = imageInfo.ImageData;
        
        if (imageInfo.HasCropping)
        {
            System.Diagnostics.Debug.WriteLine($"✂️ 检测到图片裁剪信息: 左{imageInfo.CropLeft:F1}% 上{imageInfo.CropTop:F1}% 右{imageInfo.CropRight:F1}% 下{imageInfo.CropBottom:F1}%");
            
            try
            {
                // 尝试裁剪图片数据
                finalImageData = CropImageData(imageInfo.ImageData, imageInfo.CropLeft, imageInfo.CropTop, imageInfo.CropRight, imageInfo.CropBottom);
                System.Diagnostics.Debug.WriteLine($"✂️ 图片裁剪成功，数据大小: {finalImageData.Length} bytes");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ 图片裁剪失败，使用原图: {ex.Message}");
                finalImageData = imageInfo.ImageData;
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"📷 图片无裁剪信息，使用原图");
        }
        
        // 创建图片部分，根据图片类型选择合适的ImagePart类型
        ImagePart imagePart;
        
        try
        {
            imagePart = mainPart.AddImagePart(imageInfo.ContentType);
        }
        catch
        {
            // 如果内容类型有问题，尝试使用默认的jpeg类型
            imagePart = mainPart.AddImagePart("image/jpeg");
        }
        
        // 写入处理后的图片数据
        using (var stream = new MemoryStream(finalImageData))
        {
            imagePart.FeedData(stream);
        }

        // 获取图片关系ID
        var relationshipId = mainPart.GetIdOfPart(imagePart);

        // 改进图片尺寸计算：等比例缩放适应单元格
        // 用户要求：可以等比例缩小放到单元格里
        
        // 获取图片的实际像素尺寸
        double actualWidth = 0;
        double actualHeight = 0;
        
        try
        {
            using (var stream = new MemoryStream(finalImageData))
            using (var img = Image.FromStream(stream))
            {
                actualWidth = img.Width;
                actualHeight = img.Height;
            }
        }
        catch
        {
            // 如果无法读取图片尺寸，使用默认值
            actualWidth = imageInfo.Width;
            actualHeight = imageInfo.Height;
        }
        
        // 根据OpenXML最佳实践进行尺寸计算
        // 参考：1厘米 = 360000 EMU，正确处理DPI
        
        // 从Excel图片的实际显示尺寸计算目标尺寸
        // imageInfo.Width和Height是Excel中的显示尺寸，保持比例
        double excelDisplayWidthCm = (imageInfo.Width / 72.0) * 2.54;   // 从像素转换为厘米
        double excelDisplayHeightCm = (imageInfo.Height / 72.0) * 2.54;
        
        System.Diagnostics.Debug.WriteLine($"Excel中图片显示尺寸: {excelDisplayWidthCm:F2}x{excelDisplayHeightCm:F2}cm");
        
        // 使用Excel中的实际显示尺寸作为目标，而不是固定限制
        double targetWidthCm = excelDisplayWidthCm;
        double targetHeightCm = excelDisplayHeightCm;
        
        // 直接使用Excel中的显示尺寸，确保Word中的显示与Excel一致
        double finalWidthCm = targetWidthCm;
        double finalHeightCm = targetHeightCm;
        
        // 转换为EMU（OpenXML标准单位）
        long widthEmu = (long)(finalWidthCm * 360000);   // 1cm = 360000 EMU
        long heightEmu = (long)(finalHeightCm * 360000);
        
        System.Diagnostics.Debug.WriteLine($"图片尺寸处理详情:");
        System.Diagnostics.Debug.WriteLine($"  原始图片: {actualWidth}x{actualHeight}px");
        System.Diagnostics.Debug.WriteLine($"  Excel显示: {excelDisplayWidthCm:F2}x{excelDisplayHeightCm:F2}cm");
        System.Diagnostics.Debug.WriteLine($"  Word目标: {finalWidthCm:F2}x{finalHeightCm:F2}cm");
        System.Diagnostics.Debug.WriteLine($"  EMU尺寸: {widthEmu}x{heightEmu}");

        // 创建图片元素
        var drawing = CreateImageElement(relationshipId, widthEmu, heightEmu, imageInfo.FileName);
        System.Diagnostics.Debug.WriteLine($"✅ CreateCellImageElement 完成处理图片{imageInfo.ImageIndex}");
        return drawing;
    }

    /// <summary>
    /// 裁剪图片数据
    /// </summary>
    private byte[] CropImageData(byte[] originalImageData, double cropLeft, double cropTop, double cropRight, double cropBottom)
    {
        try
        {
            using (var originalStream = new MemoryStream(originalImageData))
            using (var bitmap = Image.FromStream(originalStream))
            {
                // 计算裁剪区域
                int originalWidth = bitmap.Width;
                int originalHeight = bitmap.Height;
                
                int cropX = (int)(originalWidth * cropLeft / 100.0);
                int cropY = (int)(originalHeight * cropTop / 100.0);
                int cropWidth = (int)(originalWidth * (1.0 - (cropLeft + cropRight) / 100.0));
                int cropHeight = (int)(originalHeight * (1.0 - (cropTop + cropBottom) / 100.0));
                
                // 确保裁剪区域有效
                cropX = Math.Max(0, Math.Min(cropX, originalWidth - 1));
                cropY = Math.Max(0, Math.Min(cropY, originalHeight - 1));
                cropWidth = Math.Max(1, Math.Min(cropWidth, originalWidth - cropX));
                cropHeight = Math.Max(1, Math.Min(cropHeight, originalHeight - cropY));
                
                System.Diagnostics.Debug.WriteLine($"原图尺寸: {originalWidth}x{originalHeight}, 裁剪区域: ({cropX},{cropY}) {cropWidth}x{cropHeight}");
                
                // 如果裁剪区域太小，很可能导致图片损坏，直接返回原图
                if (cropWidth < 50 || cropHeight < 50)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ 裁剪区域过小({cropWidth}x{cropHeight})，使用原图避免损坏");
                    return originalImageData;
                }
                
                // 创建裁剪后的图片
                using (var croppedBitmap = new Bitmap(cropWidth, cropHeight))
                using (var graphics = Graphics.FromImage(croppedBitmap))
                {
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    
                    graphics.DrawImage(bitmap, 
                        new Rectangle(0, 0, cropWidth, cropHeight), 
                        new Rectangle(cropX, cropY, cropWidth, cropHeight), 
                        GraphicsUnit.Pixel);
                    
                    // 将裁剪后的图片转换为字节数组
                    using (var resultStream = new MemoryStream())
                    {
                        // 保存为JPEG格式以保持兼容性，高质量
                        var encoder = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                        var encoderParams = new EncoderParameters(1);
                        encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L);
                        
                        croppedBitmap.Save(resultStream, encoder, encoderParams);
                        byte[] result = resultStream.ToArray();
                        
                        System.Diagnostics.Debug.WriteLine($"✂️ 裁剪完成，输出大小: {result.Length} bytes");
                        
                        // 检查结果是否合理，如果太小可能有问题
                        if (result.Length < 5000) // 如果小于5KB，可能有问题
                        {
                            System.Diagnostics.Debug.WriteLine($"⚠️ 裁剪后图片过小({result.Length}bytes)，使用原图");
                            return originalImageData;
                        }
                        
                        return result;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"图片裁剪处理失败: {ex.Message}");
            throw;
        }
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

    /// <summary>
    /// 图片起始行
    /// </summary>
    public int StartRow { get; set; }

    /// <summary>
    /// 图片结束行
    /// </summary>
    public int EndRow { get; set; }

    /// <summary>
    /// 图片起始列
    /// </summary>
    public int StartCol { get; set; }

    /// <summary>
    /// 图片结束列
    /// </summary>
    public int EndCol { get; set; }

    /// <summary>
    /// 图片索引（在工作表中的顺序）
    /// </summary>
    public int ImageIndex { get; set; }

    /// <summary>
    /// 是否有裁剪
    /// </summary>
    public bool HasCropping { get; set; }

    /// <summary>
    /// 左边裁剪百分比（0-100）
    /// </summary>
    public double CropLeft { get; set; }

    /// <summary>
    /// 顶部裁剪百分比（0-100）
    /// </summary>
    public double CropTop { get; set; }

    /// <summary>
    /// 右边裁剪百分比（0-100）
    /// </summary>
    public double CropRight { get; set; }

    /// <summary>
    /// 底部裁剪百分比（0-100）
    /// </summary>
    public double CropBottom { get; set; }
}