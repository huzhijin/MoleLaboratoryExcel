using DevExpress.XtraEditors;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace MoleLaboratoryExcel.Forms
{
    public partial class ExcelToWordForm : XtraForm
    {
        private TextEdit txtFilePath;
        private SimpleButton btnBrowse;
        private SimpleButton btnExecute;
        private LabelControl lblTitle;
        private OpenFileDialog openExcelDialog;
        private string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public ExcelToWordForm()
        {
            //InitializeComponent();
            InitializeUI();
            InitializeFileDialog();
        }

        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "导出Word报告";
            this.Size = new Size(600, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 创标题标签
            lblTitle = new LabelControl
            {
                Text = "选择需生成Word报告的Excel",
                AutoSizeMode = LabelAutoSizeMode.None,
                Size = new Size(300, 20),
                Location = new Point(20, 30)
            };

            // 创建文件路径文本框
            txtFilePath = new TextEdit
            {
                Location = new Point(20, 60),
                Size = new Size(460, 20),
                ReadOnly = true,
                Properties = { NullText = "请选择Excel文件..." }
            };

            // 创建浏览按钮
            btnBrowse = new SimpleButton
            {
                Text = "...",
                Location = new Point(490, 60),
                Size = new Size(30, 20)
            };
            btnBrowse.Click += BtnBrowse_Click;

            // 创建执行按钮
            btnExecute = new SimpleButton
            {
                Text = "生成Word",
                Location = new Point(250, 150),
                Size = new Size(100, 30)
            };
            btnExecute.Click += BtnExecute_Click;

            // 添加控件到窗体
            this.Controls.AddRange(new System.Windows.Forms.Control[] {
                lblTitle,
                txtFilePath,
                btnBrowse,
                btnExecute
            });
        }

        private void InitializeFileDialog()
        {
            openExcelDialog = new OpenFileDialog
            {
                InitialDirectory = desktopPath,
                Filter = "Excel文件|*.xlsx;*.xls|所有文件|*.*",
                Multiselect = true,
                Title = "选择Excel文件"
            };
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            if (openExcelDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string filePaths = string.Join("; ", openExcelDialog.FileNames);
                    txtFilePath.Text = filePaths;
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show($"选择文件时出错：\n{ex.Message}", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnExecute_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text))
            {
                XtraMessageBox.Show("请先选择Excel文件", "提示");
                return;
            }

            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.InitialDirectory = desktopPath;
                saveDialog.DefaultExt = "docx";
                saveDialog.Filter = "Word文件|*.docx|所有文件|*.*";
                saveDialog.Title = "保存Word文件";
                
                // 设置默认文件名
                string firstExcelName = Path.GetFileNameWithoutExtension(openExcelDialog.FileNames[0]);
                saveDialog.FileName = $"{firstExcelName}_报告.docx";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ProcessExcelToWord(openExcelDialog.FileNames, saveDialog.FileName);
                        XtraMessageBox.Show("Word报告生成成功！", "提示");
                    }
                    catch (Exception ex)
                    {
                        XtraMessageBox.Show($"生成Word报告时出错：\n{ex.Message}", "错误",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ProcessExcelToWord(string[] excelFiles, string wordFile)
        {
            try
            {
                var converter = new ExcelToWordConverter();
                
                // 根据文件数量选择不同的转换方法
                if (excelFiles.Length == 1)
                {
                    // 单个Excel文件使用ConvertExcelToWord方法
                    converter.ConvertExcelToWord(excelFiles[0], wordFile);
                }
                else
                {
                    // 多个Excel文件使用ConvertMultipleExcelsToWord方法
                    converter.ConvertMultipleExcelsToWord(excelFiles, wordFile);
                }

                // 记录操作日志
                LogHelper.LogUserAction(
                    Program.CurrentUser.Username,
                    excelFiles.Length == 1 ? "SingleExcelToWord" : "MultipleExcelToWord",
                    $"生成Word报告，源文件：{string.Join(", ", excelFiles.Select(Path.GetFileName))}"
                );
            }
            catch (Exception ex)
            {
                // 记录错误日志
                LogHelper.LogError("Excel转Word失败", ex);
                throw;
            }
        }
    }
} 