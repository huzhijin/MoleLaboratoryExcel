using DevExpress.Utils.CommonDialogs;
using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Windows.Forms;

namespace MoleLaboratoryExcel.Forms
{
    public partial class ExcelImportForm : XtraForm
    {
        // 添加品牌枚举
        private enum InstrumentBrand
        {
            ThermoFisher7500, // 赛默飞7500
            HONGSHI          // 宏石
        }

        private TextEdit txtFilePath;
        private SimpleButton btnBrowse;
        private SimpleButton btnSplit;
        private SimpleButton btnMerge;
        private LabelControl lblTitle;
        private OpenFileDialog OpenExcelFileDialog;
        private ComboBoxEdit cmbBrand;
        private InstrumentBrand selectedBrand = InstrumentBrand.ThermoFisher7500;

        private List<DataTable> ExcelDataTables = new List<DataTable>();
        private Dictionary<string, DataTable> allDataDic = new Dictionary<string, DataTable>();
        private string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public ExcelImportForm()
        {
            //InitializeComponent();
            InitializeUI();
            InitializeFileDialog();
        }

        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "整理Excel";
            this.Size = new Size(500, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 创建标题标签
            lblTitle = new LabelControl
            {
                Text = "选择仪器品牌",
                AutoSizeMode = LabelAutoSizeMode.None,
                Size = new Size(200, 20),
                Location = new Point(20, 20)
            };

            // 使用ComboBoxEdit替代ComboBox
            cmbBrand = new ComboBoxEdit
            {
                Location = new Point(20, 50),
                Size = new Size(120, 25),
                Properties =
                {
                    TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor,
                    Items = { "赛默飞7500", "宏石" }
                }
            };

            // 创建文件选择标签
            var lblFile = new LabelControl
            {
                Text = "选择需要整合的Excel",
                AutoSizeMode = LabelAutoSizeMode.None,
                Size = new Size(200, 20),
                Location = new Point(20, 90)
            };

            // 创建文件路径文本框
            txtFilePath = new TextEdit
            {
                Location = new Point(20, 120),
                Size = new Size(360, 20),
                ReadOnly = true,
                Properties = { NullText = "请选择Excel文件..." }
            };

            // 创建浏览按钮
            btnBrowse = new SimpleButton
            {
                Text = "浏览",
                Location = new Point(390, 120),
                Size = new Size(60, 20)
            };
            btnBrowse.Click += BtnBrowse_Click;

            // 创建分开整理按钮
            btnSplit = new SimpleButton
            {
                Text = "分开整理",
                Location = new Point(120, 180),
                Size = new Size(100, 30)
            };
            btnSplit.Click += BtnSplit_Click;

            // 创建合并整理按钮
            btnMerge = new SimpleButton
            {
                Text = "合并整理",
                Location = new Point(280, 180),
                Size = new Size(100, 30)
            };
            btnMerge.Click += BtnMerge_Click;

            // 设置默认选中项
            cmbBrand.SelectedIndex = 0;
            cmbBrand.SelectedIndexChanged += CmbBrand_SelectedIndexChanged;

            // 添加到控件列表
            this.Controls.AddRange(new Control[] {
                lblTitle,
                cmbBrand,
                lblFile,
                txtFilePath,
                btnBrowse,
                btnSplit,
                btnMerge
            });
        }

        private void InitializeFileDialog()
        {
            OpenExcelFileDialog = new OpenFileDialog
            {
                InitialDirectory = desktopPath,
                Filter = "Excel文件|*.xlsx;*.xls|所有文件|*.*",
                Multiselect = true,
                Title = "选择文件"
            };
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            if (OpenExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = "";
                try
                {
                    string filePath = "";
                    ExcelDataTables.Clear();
                    allDataDic.Clear();

                    foreach (string file in OpenExcelFileDialog.FileNames)
                    {
                        string filename = Path.GetFileName(file);
                        // 根据不同品牌设置不同的起始行
                        int headerRowIndex = selectedBrand == InstrumentBrand.HONGSHI ? 13 : 7; // 13对应第14行
                        DataTable excelDataTable = DataTableUtil.ExcelToDataTable(file, headerRowIndex);

                        // 根据品牌重命名列
                        //if (selectedBrand == InstrumentBrand.HONGSHI)
                        //{
                        //    RenameColumns(excelDataTable);
                        //}

                        ExcelDataTables.Add(excelDataTable);
                        filePath += file + "; ";
                        allDataDic.Add(filename, excelDataTable);
                    }
                    txtFilePath.Text = filePath.TrimEnd(';', ' ');

                    // 记录导入操作日志
                    LogHelper.LogUserAction(
                        Program.CurrentUser.Username,
                        "ImportExcel",
                        $"导入Excel文件：{string.Join(", ", OpenExcelFileDialog.FileNames.Select(Path.GetFileName))}"
                    );
                }
                catch (Exception ex)
                {
                    LogHelper.LogError("导入Excel失败", ex);
                    XtraMessageBox.Show($"访问文件权限错误：\n{ex.Message}", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void RenameColumns(DataTable table)
        {
            var columnMappings = new Dictionary<string, string>
            {
                { "反应孔", "Well" },
                { "样本名称", "Sample Name" },
                { "目标", "Target Name" },
                { "Ct", "Cт" },
                { "属性", "Quantity" }
            };

            foreach (var mapping in columnMappings)
            {
                if (table.Columns.Contains(mapping.Key))
                {
                    table.Columns[mapping.Key].ColumnName = mapping.Value;
                }
            }
        }

        private void BtnMerge_Click(object sender, EventArgs e)
        {
            if (ExcelDataTables.Count == 0)
            {
                XtraMessageBox.Show("请先选择Excel文件", "提示");
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.InitialDirectory = desktopPath;
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.Filter = "Excel文件|*.xlsx|所有文件|*.*";
                saveFileDialog.Title = "保存合并文件";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var processedTables = new Dictionary<string, DataTable>();
                        foreach (var kvp in allDataDic)
                        {
                            string[] columnNames = GetColumnNames();
                            DataTable processedTable = DataTableUtil.RetainColumns(kvp.Value, columnNames);
                            DataTable transformedTable;

                            if (selectedBrand == InstrumentBrand.ThermoFisher7500)
                            {
                                processedTable = DataTableUtil.RemoveEmptyOrNullRowsEfficiently(
                                    processedTable, "Target Name");
                                transformedTable = DataTableUtil.TransformQPCRData(processedTable);
                                DataTableUtil.AddWellColumns(transformedTable, "Well");
                            }
                            else
                            {
                                processedTable = DataTableUtil.RemoveEmptyOrNullRowsEfficiently(
                                    processedTable, "目标");
                                transformedTable = DataTableUtil.TransformRowsToColumns(
                                    processedTable,
                                    groupByColumns: new[] { "反应孔", "样本名称" },
                                    categoryColumn: "目标",
                                    valueColumns: new[] { "Ct", "属性" },
                                    suffixes: new[] { "-Ct", "-属性" }
                                );
                                DataTableUtil.AddWellColumns(transformedTable, "反应孔");
                            }

                            processedTables.Add(kvp.Key, transformedTable);
                        }

                        // 将所有表格写入到一个Excel文件的不同sheet中
                        DataTableUtil.WriteTableToExcelSheets(processedTables, saveFileDialog.FileName);

                        // 记录合并操作日志
                        LogHelper.LogUserAction(
                            Program.CurrentUser.Username,
                            "MergeExcel",
                            $"合并Excel文件，保存为：{Path.GetFileName(saveFileDialog.FileName)}"
                        );

                        XtraMessageBox.Show("合并整理完成！", "提示");
                    }
                    catch (Exception ex)
                    {
                        // 记录错误日志
                        LogHelper.LogError("合并Excel失败", ex);
                        XtraMessageBox.Show($"合并处理出错：\n{ex.Message}", "错误",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnSplit_Click(object sender, EventArgs e)
        {
            if (ExcelDataTables.Count == 0)
            {
                XtraMessageBox.Show("请先选择Excel文件", "提示");
                return;
            }

            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "选择保存位置";
                folderDialog.SelectedPath = desktopPath;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        List<string> processedFiles = new List<string>();
                        foreach (var kvp in allDataDic)
                        {
                            string[] columnNames = GetColumnNames();
                            DataTable processedTable = DataTableUtil.RetainColumns(kvp.Value, columnNames);
                            DataTable transformedTable = new DataTable();

                            if (selectedBrand == InstrumentBrand.ThermoFisher7500)
                            {
                                processedTable = DataTableUtil.RemoveEmptyOrNullRowsEfficiently(
                              processedTable, "Target Name");
                                transformedTable = DataTableUtil.TransformQPCRData(processedTable);
                                DataTableUtil.AddWellColumns(transformedTable, "Well");
                            }
                            else
                            {
                                processedTable = DataTableUtil.RemoveEmptyOrNullRowsEfficiently(
                               processedTable, "目标");
                                transformedTable = DataTableUtil.TransformRowsToColumns(
                   processedTable,
                   groupByColumns: new[] { "反应孔", "样本名称" },             // 按Well分��
                   categoryColumn: "目标",       // 使用Target Name作为类别
                   valueColumns: new[] { "Ct", "属性" },  // 需要转换的值列
                   suffixes: new[] { "-Ct", "-属性" }    // 对应的后缀
                    
               );
                                DataTableUtil.AddWellColumns(transformedTable, "反应孔");
                            }

                            string saveFilename = "整理后_" + kvp.Key;
                            string saveFilePath = Path.Combine(folderDialog.SelectedPath, saveFilename);

                            DataTableUtil.WriteTableToExcel(transformedTable, saveFilePath);
                            processedFiles.Add(saveFilename);
                        }

                        // 记录分开整理操作日志
                        LogHelper.LogUserAction(
                            Program.CurrentUser.Username,
                            "SplitExcel",
                            $"分开整理Excel文件，生成文件：{string.Join(", ", processedFiles)}"
                        );

                        XtraMessageBox.Show("分开整理完成！", "提示");
                    }
                    catch (Exception ex)
                    {
                        // 记录错误日志
                        LogHelper.LogError("分开整理Excel失败", ex);
                        XtraMessageBox.Show($"整理过程中出错：\n{ex.Message}", "错误",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void CmbBrand_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedBrand = cmbBrand.SelectedIndex == 0 ?
                InstrumentBrand.ThermoFisher7500 :
                InstrumentBrand.HONGSHI;
        }

        private string[] GetColumnNames()
        {
            // 根据不同品牌返回对应列名数组
            return selectedBrand == InstrumentBrand.ThermoFisher7500
                ? new[] { "Well", "Sample Name", "Target Name", "Cт", "Quantity" }
                : new[] { "反应孔", "样本名称", "目标", "Ct", "属性" };
        }
    }
}