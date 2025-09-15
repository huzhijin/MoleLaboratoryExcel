using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace MoleLaboratoryExcel.Forms
{
    public partial class BatchRenameForm : XtraForm
    {
        // Windows API for natural sorting
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);

        private TextEdit txtTxtFilePath;
        private TextEdit txtFolderPath;
        private SimpleButton btnBrowseTxt;
        private SimpleButton btnBrowseFolder;
        private SimpleButton btnRename;
        private LabelControl lblTxtFile;
        private LabelControl lblFolder;
        private LabelControl lblTitle;
        private OpenFileDialog openTxtFileDialog;
        private FolderBrowserDialog folderBrowserDialog;

        public BatchRenameForm()
        {
            InitializeUI();
            InitializeDialogs();
        }

        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "批量重命名文件";
            this.Size = new Size(500, 280);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 创建标题标签
            lblTitle = new LabelControl
            {
                Text = "批量重命名文件",
                Font = new Font("微软雅黑", 12, FontStyle.Bold),
                AutoSizeMode = LabelAutoSizeMode.None,
                Size = new Size(200, 25),
                Location = new Point(20, 20)
            };

            // 创建txt文件选择标签
            lblTxtFile = new LabelControl
            {
                Text = "选择名称文件（txt）:",
                AutoSizeMode = LabelAutoSizeMode.None,
                Size = new Size(150, 20),
                Location = new Point(20, 60)
            };

            // 创建txt文件路径文本框
            txtTxtFilePath = new TextEdit
            {
                Location = new Point(20, 85),
                Size = new Size(360, 20),
                ReadOnly = true,
                Properties = { NullText = "请选择包含文件名的txt文件..." }
            };

            // 创建txt文件浏览按钮
            btnBrowseTxt = new SimpleButton
            {
                Text = "浏览",
                Location = new Point(390, 85),
                Size = new Size(60, 20)
            };
            btnBrowseTxt.Click += BtnBrowseTxt_Click;

            // 创建文件夹选择标签
            lblFolder = new LabelControl
            {
                Text = "选择要重命名的文件夹:",
                AutoSizeMode = LabelAutoSizeMode.None,
                Size = new Size(150, 20),
                Location = new Point(20, 125)
            };

            // 创建文件夹路径文本框
            txtFolderPath = new TextEdit
            {
                Location = new Point(20, 150),
                Size = new Size(360, 20),
                ReadOnly = true,
                Properties = { NullText = "请选择包含待重命名文件的文件夹..." }
            };

            // 创建文件夹浏览按钮
            btnBrowseFolder = new SimpleButton
            {
                Text = "浏览",
                Location = new Point(390, 150),
                Size = new Size(60, 20)
            };
            btnBrowseFolder.Click += BtnBrowseFolder_Click;

            // 创建重命名按钮
            btnRename = new SimpleButton
            {
                Text = "执行重命名",
                Location = new Point(200, 200),
                Size = new Size(100, 30),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            btnRename.Click += BtnRename_Click;

            // 添加到控件列表
            this.Controls.AddRange(new Control[] {
                lblTitle,
                lblTxtFile,
                txtTxtFilePath,
                btnBrowseTxt,
                lblFolder,
                txtFolderPath,
                btnBrowseFolder,
                btnRename
            });
        }

        private void InitializeDialogs()
        {
            // 初始化txt文件选择对话框
            openTxtFileDialog = new OpenFileDialog
            {
                Filter = "文本文件|*.txt|所有文件|*.*",
                Title = "选择包含文件名的txt文件",
                Multiselect = false
            };

            // 初始化文件夹选择对话框
            folderBrowserDialog = new FolderBrowserDialog
            {
                Description = "选择包含待重命名文件的文件夹"
            };
        }

        private void BtnBrowseTxt_Click(object sender, EventArgs e)
        {
            if (openTxtFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtTxtFilePath.Text = openTxtFileDialog.FileName;
            }
        }

        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                txtFolderPath.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private void BtnRename_Click(object sender, EventArgs e)
        {
            try
            {
                // 验证输入
                if (string.IsNullOrEmpty(txtTxtFilePath.Text))
                {
                    XtraMessageBox.Show("请选择包含文件名的txt文件!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(txtFolderPath.Text))
                {
                    XtraMessageBox.Show("请选择要重命名的文件夹!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查文件和文件夹是否存在
                if (!File.Exists(txtTxtFilePath.Text))
                {
                    XtraMessageBox.Show("选择的txt文件不存在!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!Directory.Exists(txtFolderPath.Text))
                {
                    XtraMessageBox.Show("选择的文件夹不存在!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 读取txt文件内容（支持多种编码）
                string[] newFileNames = ReadTextFileWithEncoding(txtTxtFilePath.Text);

                // 获取文件夹中的文件并排序
                string[] existingFiles = GetSortedFiles(txtFolderPath.Text);

                // 检查数量是否一致
                if (newFileNames.Length != existingFiles.Length)
                {
                    XtraMessageBox.Show(
                        $"文件数量不匹配!\n" +
                        $"txt文件中有 {newFileNames.Length} 行文件名\n" +
                        $"文件夹中有 {existingFiles.Length} 个文件",
                        "错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // 确认重命名操作
                var result = XtraMessageBox.Show(
                    $"即将重命名 {existingFiles.Length} 个文件，确定继续吗？",
                    "确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                {
                    return;
                }

                // 执行重命名
                int successCount = 0;
                string renameDetails = "";

                for (int i = 0; i < existingFiles.Length; i++)
                {
                    try
                    {
                        string oldFilePath = existingFiles[i];
                        string oldFileName = Path.GetFileName(oldFilePath);
                        string extension = Path.GetExtension(oldFilePath);

                        // 构建新文件名，确保编码正确
                        string baseNewFileName = newFileNames[i];
                        if (string.IsNullOrWhiteSpace(baseNewFileName))
                        {
                            continue; // 跳过空的文件名
                        }

                        string newFileName = baseNewFileName + extension;
                        string newFilePath = Path.Combine(txtFolderPath.Text, newFileName);

                        // 如果新文件名和旧文件名相同，跳过
                        if (string.Equals(oldFileName, newFileName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        // 验证新文件名是否有效
                        if (!IsValidFileName(newFileName))
                        {
                            XtraMessageBox.Show(
                                $"无效的文件名: {newFileName}\n请检查文件名是否包含非法字符！",
                                "错误",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return;
                        }

                        // 检查目标文件是否已存在
                        if (File.Exists(newFilePath))
                        {
                            XtraMessageBox.Show(
                                $"目标文件已存在: {newFileName}\n重命名操作已停止！",
                                "错误",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return;
                        }

                        // 执行重命名
                        File.Move(oldFilePath, newFilePath);
                        successCount++;
                        renameDetails += $"{oldFileName} → {newFileName}\n";

                        // 验证重命名是否成功
                        if (!File.Exists(newFilePath))
                        {
                            throw new Exception("重命名后文件不存在，可能是编码问题");
                        }
                    }
                    catch (Exception ex)
                    {
                        XtraMessageBox.Show(
                            $"重命名文件失败: {Path.GetFileName(existingFiles[i])}\n错误: {ex.Message}",
                            "错误",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }
                }

                // 记录操作日志
                LogHelper.LogUserAction(
                    Program.CurrentUser.Username,
                    "BatchRename",
                    $"批量重命名文件，成功重命名 {successCount} 个文件"
                );

                // 显示成功消息
                XtraMessageBox.Show(
                    $"批量重命名完成！\n成功重命名 {successCount} 个文件。",
                    "成功",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // 清空输入框
                txtTxtFilePath.Text = "";
                txtFolderPath.Text = "";
            }
            catch (Exception ex)
            {
                LogHelper.LogError("批量重命名失败", ex);
                XtraMessageBox.Show(
                    $"批量重命名过程中发生错误:\n{ex.Message}",
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private string[] ReadTextFileWithEncoding(string filePath)
        {
            try
            {
                // 检测文件编码并读取
                byte[] fileBytes = File.ReadAllBytes(filePath);
                string content = DetectEncodingAndReadText(fileBytes);

                // 分割行并处理每行
                return content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => NormalizeFileName(line.Trim()))
                    .Where(line => !string.IsNullOrWhiteSpace(line))
                    .ToArray();
            }
            catch (Exception ex)
            {
                throw new Exception($"读取文件失败: {ex.Message}", ex);
            }
        }

        private string DetectEncodingAndReadText(byte[] fileBytes)
        {
            // 检测BOM
            if (fileBytes.Length >= 3 && fileBytes[0] == 0xEF && fileBytes[1] == 0xBB && fileBytes[2] == 0xBF)
            {
                // UTF-8 BOM
                return Encoding.UTF8.GetString(fileBytes, 3, fileBytes.Length - 3);
            }
            else if (fileBytes.Length >= 2 && fileBytes[0] == 0xFF && fileBytes[1] == 0xFE)
            {
                // UTF-16 LE BOM
                return Encoding.Unicode.GetString(fileBytes, 2, fileBytes.Length - 2);
            }
            else if (fileBytes.Length >= 2 && fileBytes[0] == 0xFE && fileBytes[1] == 0xFF)
            {
                // UTF-16 BE BOM
                return Encoding.BigEndianUnicode.GetString(fileBytes, 2, fileBytes.Length - 2);
            }

            // 尝试不同编码
            string[] encodingsToTry = { "UTF-8", "GBK", "GB2312", "Big5" };

            foreach (string encodingName in encodingsToTry)
            {
                try
                {
                    Encoding encoding = Encoding.GetEncoding(encodingName);
                    string result = encoding.GetString(fileBytes);

                    // 简单验证：检查是否包含过多的替换字符
                    if (result.Count(c => c == '\uFFFD') < result.Length * 0.1) // 替换字符少于10%
                    {
                        return result;
                    }
                }
                catch
                {
                    continue;
                }
            }

            // 最后使用系统默认编码
            return Encoding.Default.GetString(fileBytes);
        }

        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return fileName;

            // 移除或替换Windows文件名中的非法字符
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string normalized = fileName;

            foreach (char invalidChar in invalidChars)
            {
                normalized = normalized.Replace(invalidChar, '_');
            }

            // 移除首尾空格和点号
            normalized = normalized.Trim(' ', '.');

            // 处理保留名称
            string[] reservedNames = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
            string upperNormalized = normalized.ToUpper();

            if (reservedNames.Contains(upperNormalized))
            {
                normalized = "_" + normalized;
            }

            // 限制文件名长度（Windows文件名最大255字符，但考虑扩展名，这里限制为200）
            if (normalized.Length > 200)
            {
                normalized = normalized.Substring(0, 200);
            }

            return normalized;
        }

        private bool IsValidFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return false;

            // 检查是否包含非法字符
            char[] invalidChars = Path.GetInvalidFileNameChars();
            if (fileName.IndexOfAny(invalidChars) >= 0)
                return false;

            // 检查长度
            if (fileName.Length > 255)
                return false;

            // 检查是否为保留名称
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName).ToUpper();
            string[] reservedNames = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };

            if (reservedNames.Contains(nameWithoutExtension))
                return false;

            // 检查是否以点号或空格结尾
            if (fileName.EndsWith(".") || fileName.EndsWith(" "))
                return false;

            return true;
        }

        private string[] GetSortedFiles(string folderPath)
        {
            var files = Directory.GetFiles(folderPath);

            // 询问用户选择排序方式
            var sortChoice = XtraMessageBox.Show(
                "请选择文件排序方式：\n\n" +
                "是(Y) - 按修改时间排序（推荐，通常与文件创建顺序一致）\n" +
                "否(N) - 按文件名自然排序\n" +
                "取消 - 按创建时间排序",
                "选择排序方式",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            try
            {
                switch (sortChoice)
                {
                    case DialogResult.Yes:
                        // 按修改时间排序
                        return files.OrderBy(f => File.GetLastWriteTime(f)).ToArray();

                    case DialogResult.No:
                        // 按文件名自然排序
                        return files.OrderBy(f => Path.GetFileName(f), new NaturalStringComparer()).ToArray();

                    case DialogResult.Cancel:
                    default:
                        // 按创建时间排序
                        return files.OrderBy(f => File.GetCreationTime(f)).ToArray();
                }
            }
            catch (Exception ex)
            {
                // 如果出错，使用默认的文件名排序
                XtraMessageBox.Show($"排序时出错，使用默认排序: {ex.Message}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return files.OrderBy(f => Path.GetFileName(f)).ToArray();
            }
        }
    }

    // 跨平台自然排序比较器
    public class NaturalStringComparer : IComparer<string>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);

        public int Compare(string x, string y)
        {
            // 在Windows上使用系统API
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                try
                {
                    return StrCmpLogicalW(x, y);
                }
                catch
                {
                    // 如果API调用失败，回退到自定义实现
                    return CompareNatural(x, y);
                }
            }
            else
            {
                // 在非Windows平台使用自定义实现
                return CompareNatural(x, y);
            }
        }

        private int CompareNatural(string x, string y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            int i = 0, j = 0;
            while (i < x.Length && j < y.Length)
            {
                if (char.IsDigit(x[i]) && char.IsDigit(y[j]))
                {
                    // 比较数字部分
                    string numX = "";
                    string numY = "";

                    while (i < x.Length && char.IsDigit(x[i]))
                        numX += x[i++];

                    while (j < y.Length && char.IsDigit(y[j]))
                        numY += y[j++];

                    if (int.TryParse(numX, out int intX) && int.TryParse(numY, out int intY))
                    {
                        int result = intX.CompareTo(intY);
                        if (result != 0) return result;
                    }
                    else
                    {
                        int result = numX.CompareTo(numY);
                        if (result != 0) return result;
                    }
                }
                else
                {
                    // 比较字符部分
                    int result = char.ToLowerInvariant(x[i]).CompareTo(char.ToLowerInvariant(y[j]));
                    if (result != 0) return result;
                    i++;
                    j++;
                }
            }

            return x.Length.CompareTo(y.Length);
        }
    }
}