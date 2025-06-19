using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using MoleLaboratoryExcel;
using System;
using System.Windows.Forms;

public partial class LogQueryForm : XtraForm
{
    private GridControl gridControl;
    private GridView gridView;
    private DateEdit dateStart;
    private DateEdit dateEnd;
    private TextEdit txtUsername;
    private ComboBoxEdit cmbAction;
    private SimpleButton btnQuery;
    private SimpleButton btnExport;

    public LogQueryForm()
    {
        //InitializeComponent();
        InitializeUI();
        // 注册窗体加载事件
        this.Load += LogQueryForm_Load;
    }

    private void LogQueryForm_Load(object sender, EventArgs e)
    {
        try
        {
            // 执行默认查询
            QueryLogs();
            LogHelper.LogUserAction(Program.CurrentUser.Username, "查询日志", "打开日志查询界面时自动查询");
        }
        catch (Exception ex)
        {
            XtraMessageBox.Show("查询日志数据失败：" + ex.Message, "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            LogHelper.LogError("日志查询窗体加载查询失败", ex);
        }
    }

    private void QueryLogs()
    {
        // 调用查询按钮的点击事件
        BtnQuery_Click(this, EventArgs.Empty);
    }

    private void InitializeUI()
    {
        this.Text = "操作日志查询";
        this.Size = new System.Drawing.Size(1000, 600);
        this.StartPosition = FormStartPosition.CenterScreen;

        // 创建主布局面板
        var mainPanel = new PanelControl
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(10)
        };

        // 创建查询面板
        var panelQuery = new PanelControl
        {
            Dock = DockStyle.Top,
            Height = 50,  // 减小高度
            Padding = new Padding(5)
        };

        // 开始时间
        var lblStart = new LabelControl
        {
            Text = "开始时间：",
            Location = new System.Drawing.Point(10, 10)
        };

        dateStart = new DateEdit
        {
            Location = new System.Drawing.Point(70, 8),
            EditValue = DateTime.Today
        };

        // 结束时间
        var lblEnd = new LabelControl
        {
            Text = "结束时间：",
            Location = new System.Drawing.Point(220, 10)
        };

        dateEnd = new DateEdit
        {
            Location = new System.Drawing.Point(280, 8),
            EditValue = DateTime.Now
        };

        // 用户名
        var lblUsername = new LabelControl
        {
            Text = "用户名：",
            Location = new System.Drawing.Point(430, 10)
        };

        txtUsername = new TextEdit
        {
            Location = new System.Drawing.Point(480, 8),
            Size = new System.Drawing.Size(100, 20)
        };

        // 操作类型
        var lblAction = new LabelControl
        {
            Text = "操作类型：",
            Location = new System.Drawing.Point(590, 10)
        };

        cmbAction = new ComboBoxEdit
        {
            Location = new System.Drawing.Point(650, 8),
            Size = new System.Drawing.Size(100, 20)
        };
        cmbAction.Properties.Items.AddRange(new[] {
            "全部",
            "登录",
            "新增用户",
            "修改用户",
            "删除用户",
            "导入Excel",
            "导出Word",
            "错误"
        });
        cmbAction.SelectedIndex = 0;  // 默认选择"全部"

        // 查询按钮
        btnQuery = new SimpleButton
        {
            Text = "查询",
            Location = new System.Drawing.Point(760, 7)
        };
        btnQuery.Click += BtnQuery_Click;

        // 导出按钮
        btnExport = new SimpleButton
        {
            Text = "导出",
            Location = new System.Drawing.Point(850, 7)
        };
        btnExport.Click += BtnExport_Click;

        // 添加控件到查询面板
        panelQuery.Controls.AddRange(new Control[]
        {
            lblStart, dateStart,
            lblEnd, dateEnd,
            lblUsername, txtUsername,
            lblAction, cmbAction,
            btnQuery, btnExport
        });

        // 创建表格容器面板
        var gridPanel = new PanelControl
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 5, 0, 0)  // 只添加顶部间距
        };

        // 修改表格控件设置
        gridControl = new GridControl
        {
            Dock = DockStyle.Fill
        };

        gridView = new GridView(gridControl);
        gridControl.MainView = gridView;

        // 配置表格列
        var colId = gridView.Columns.AddVisible("Id", "ID");
        colId.Visible = false;  // 隐藏ID列

        var colUserId = gridView.Columns.AddVisible("UserId", "User ID");
        colUserId.Visible = false;  // 隐藏User ID列

        gridView.Columns.AddVisible("Username", "用户名");
        gridView.Columns.AddVisible("Action", "操作类型");
        gridView.Columns.AddVisible("Description", "操作描述");
        gridView.Columns.AddVisible("LogTime", "操作时间");
        gridView.Columns.AddVisible("IPAddress", "IP地址");

        // 设置列宽和其他属性
        foreach (DevExpress.XtraGrid.Columns.GridColumn col in gridView.Columns)
        {
            switch (col.FieldName)
            {
                case "Username":
                    col.Width = 100;
                    break;
                case "Action":
                    col.Width = 100;
                    break;
                case "Description":
                    col.Width = 200;
                    break;
                case "LogTime":
                    col.Width = 150;
                    col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    col.DisplayFormat.FormatString = "yyyy-MM-dd HH:mm:ss";
                    break;
                case "IPAddress":
                    col.Width = 120;
                    break;
            }
        }

        // 设置表格属性
        gridView.OptionsSelection.EnableAppearanceFocusedCell = false;
        gridView.OptionsBehavior.Editable = false;
        gridView.OptionsView.ShowGroupPanel = false;
        gridView.OptionsView.ShowIndicator = false;  // 隐藏行号
        gridView.OptionsView.EnableAppearanceEvenRow = true;  // 启用偶数行背景色
        gridView.OptionsView.ShowHorizontalLines = DevExpress.Utils.DefaultBoolean.True;
        gridView.OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.True;

        // 添加控件到容器
        gridPanel.Controls.Add(gridControl);
        mainPanel.Controls.Add(gridPanel);
        mainPanel.Controls.Add(panelQuery);

        // 添加到窗体
        this.Controls.Add(mainPanel);
    }

    private void LoadData()
    {
        try
        {
            var logDao = new LogDao();
            var logs = logDao.GetLogs(
                dateStart.DateTime,
                dateEnd.DateTime,
                txtUsername.Text,
                cmbAction.Text
            );
            gridControl.DataSource = logs;
        }
        catch (Exception ex)
        {
            XtraMessageBox.Show("加载日志数据失败：" + ex.Message, "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnQuery_Click(object sender, EventArgs e)
    {
        try
        {
            var logDao = new LogDao();

            // 获取操作类型的实际值
            string action = null;
            if (cmbAction.Text != "全部")
            {
                // 将中文操作类型转换为英文
                switch (cmbAction.Text)
                {
                    case "登录":
                        action = "Login";
                        break;
                    case "新增用户":
                        action = "AddUser";
                        break;
                    case "修改用户":
                        action = "UpdateUser";
                        break;
                    case "删除用户":
                        action = "DeleteUser";
                        break;
                    case "导入Excel":
                        action = "ImportExcel";
                        break;
                    case "导出Word":
                        action = "ExportWord";
                        break;
                    case "错误":
                        action = "Error";
                        break;
                }
            }

            var logs = logDao.GetLogs(
                dateStart.DateTime,
                dateEnd.DateTime.AddDays(1).AddSeconds(-1),  // 设置为当天的最后一秒
                txtUsername.Text,
                action
            );
            gridControl.DataSource = logs;
        }
        catch (Exception ex)
        {
            XtraMessageBox.Show("查询失败：" + ex.Message, "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnExport_Click(object sender, EventArgs e)
    {
        try
        {
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel文件|*.xlsx";
                saveDialog.Title = "导出日志";
                saveDialog.FileName = $"操作日志_{DateTime.Now:yyyyMMddHHmmss}";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    gridControl.ExportToXlsx(saveDialog.FileName);
                    XtraMessageBox.Show("导出成功！", "提示");
                }
            }
        }
        catch (Exception ex)
        {
            XtraMessageBox.Show("导出失败：" + ex.Message, "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
} 