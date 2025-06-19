using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Windows.Forms;
using System.Drawing;


namespace MoleLaboratoryExcel.Forms
{
    public partial class UserManageForm : XtraForm
    {
        private GridControl gridControl;
        private GridView gridView;
        private PanelControl panelTop;
        private SimpleButton btnAdd;
        private SimpleButton btnEdit;
        private SimpleButton btnDelete;
        private SimpleButton btnRefresh;
        private GroupControl editGroup;
        private TextEdit txtUsername;
        private TextEdit txtPassword;
        private ComboBoxEdit cmbRole;
        private CheckEdit chkIsActive;
        private SimpleButton btnSave;
        private SimpleButton btnCancel;

        private int currentUserId = 0;

        public UserManageForm()
        {
            //InitializeComponent();
            InitializeUI();
            LoadData();
            // 注册窗体加载事件
            this.Load += UserManageForm_Load;
        }

        private void UserManageForm_Load(object sender, EventArgs e)
        {
            try
            {
                // 执行默认查询
                QueryUsers();
                LogHelper.LogUserAction(Program.CurrentUser.Username, "查询用户", "打开用户管理界面时自动查询");
            }
            catch (Exception ex)
            {
                MessageBox.Show("查询用户数据失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogHelper.LogError("用户管理窗体加载查询失败", ex);
            }
        }

        private void QueryUsers()
        {
            LoadData();
        }

        private void InitializeUI()
        {
            this.Text = "用户管理";
            this.Size = new System.Drawing.Size(1000, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // 创建主布局面板
            var mainPanel = new PanelControl
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // 创建工具栏面板
            panelTop = new PanelControl
            {
                Dock = DockStyle.Top,
                Height = 50,  // 减小高度
                Padding = new Padding(5)
            };

            // 添加工具栏按钮
            btnAdd = new SimpleButton
            {
                Text = "新增",
                Location = new System.Drawing.Point(10, 10),
                Width = 80
            };
            btnAdd.Click += BtnAdd_Click;

            btnEdit = new SimpleButton
            {
                Text = "编辑",
                Location = new System.Drawing.Point(100, 10),
                Width = 80
            };
            btnEdit.Click += BtnEdit_Click;

            btnDelete = new SimpleButton
            {
                Text = "删除",
                Location = new System.Drawing.Point(190, 10),
                Width = 80
            };
            btnDelete.Click += BtnDelete_Click;

            btnRefresh = new SimpleButton
            {
                Text = "刷新",
                Location = new System.Drawing.Point(280, 10),
                Width = 80
            };
            btnRefresh.Click += BtnRefresh_Click;

            panelTop.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDelete, btnRefresh });

            // 创建内容面板
            var contentPanel = new PanelControl
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 5, 0, 0)
            };

            // 创建右侧编辑区域
            editGroup = new GroupControl
            {
                Text = "用户信息",
                Dock = DockStyle.Right,
                Width = 300,
                Visible = false  // 初始时隐藏
            };

            // 创建编辑区域的控件
            var lblUsername = new LabelControl
            {
                Text = "用户名：",
                Location = new Point(20, 30)
            };

            txtUsername = new TextEdit
            {
                Location = new Point(80, 27),
                Width = 180
            };

            var lblPassword = new LabelControl
            {
                Text = "密码：",
                Location = new Point(20, 60)
            };

            txtPassword = new TextEdit
            {
                Location = new Point(80, 57),
                Width = 180,
                Properties = { PasswordChar = '*' }
            };

            var lblRole = new LabelControl
            {
                Text = "角色：",
                Location = new Point(20, 90)
            };

            cmbRole = new ComboBoxEdit
            {
                Location = new Point(80, 87),
                Width = 180
            };
            cmbRole.Properties.Items.AddRange(new[] { "管理员", "普通用户" });

            chkIsActive = new CheckEdit
            {
                Text = "启用账号",
                Location = new Point(80, 117),
                Checked = true
            };

            btnSave = new SimpleButton
            {
                Text = "保存",
                Location = new Point(80, 150),
                Width = 80
            };
            btnSave.Click += BtnSave_Click;

            btnCancel = new SimpleButton
            {
                Text = "取消",
                Location = new Point(180, 150),
                Width = 80
            };
            btnCancel.Click += BtnCancel_Click;

            // 添加控件到编辑组
            editGroup.Controls.AddRange(new Control[] {
                lblUsername, txtUsername,
                lblPassword, txtPassword,
                lblRole, cmbRole,
                chkIsActive,
                btnSave, btnCancel
            });

            // 创建表格控件
            gridControl = new GridControl
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };

            gridView = new GridView(gridControl);
            gridControl.MainView = gridView;

            // 设置表格视图的选项
            gridView.OptionsView.ShowGroupPanel = false;
            gridView.OptionsView.GroupFooterShowMode = GroupFooterShowMode.Hidden;
            gridView.RowHeight = 25;
            gridView.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True;

            // 优化表格显示
            gridView.OptionsView.ShowIndicator = false;  // 隐藏行号
            gridView.OptionsView.ShowHorizontalLines = DevExpress.Utils.DefaultBoolean.True;
            gridView.OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.True;
            gridView.OptionsView.EnableAppearanceEvenRow = true;  // 启用偶数行背景色

            // 配置表格列
            var colId = gridView.Columns.AddVisible("Id", "ID");
            colId.Visible = false;  // 隐藏ID列

            gridView.Columns.AddVisible("Username", "用户名");
            gridView.Columns.AddVisible("Role", "角色");
            gridView.Columns.AddVisible("CreateTime", "创建时间");
            gridView.Columns.AddVisible("IsActive", "是否启用");

            // 设置列属性
            gridView.OptionsSelection.EnableAppearanceFocusedCell = false;
            gridView.OptionsBehavior.Editable = false;

            // 添加控件到容器
            contentPanel.Controls.Add(gridControl);
            contentPanel.Controls.Add(editGroup);
            mainPanel.Controls.Add(contentPanel);
            mainPanel.Controls.Add(panelTop);

            // 添加到窗体
            this.Controls.Add(mainPanel);

            // 初始化编辑模式
            SetEditMode(false);
        }

        private void LoadData()
        {
            try
            {
                var userDao = new UserDao();
                gridControl.DataSource = userDao.GetAllUsers();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("加载用户数据失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetEditMode(bool isEdit)
        {
            editGroup.Enabled = isEdit;
            gridControl.Enabled = !isEdit;
            btnAdd.Enabled = !isEdit;
            btnEdit.Enabled = !isEdit;
            btnDelete.Enabled = !isEdit;
        }

        private void ClearInputs()
        {
            currentUserId = 0;
            txtUsername.Text = "";
            txtPassword.Text = "";
            cmbRole.SelectedIndex = -1;
            chkIsActive.Checked = true;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            ClearInputs();
            SetEditMode(true);
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            var row = gridView.GetFocusedRow() as User;
            if (row == null)
            {
                XtraMessageBox.Show("请选择要编辑的用户", "提示");
                return;
            }

            currentUserId = row.Id;
            txtUsername.Text = row.Username;
            txtPassword.Text = "";  // 不显示密码
            cmbRole.Text = row.Role;
            chkIsActive.Checked = row.IsActive;

            SetEditMode(true);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            var row = gridView.GetFocusedRow() as User;
            if (row == null)
            {
                XtraMessageBox.Show("请选择要删除的用户", "提示");
                return;
            }

            if (XtraMessageBox.Show("确定要删除该用户吗？", "确认",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var userDao = new UserDao();
                    if (userDao.DeleteUser(row.Id))
                    {
                        LogHelper.LogUserAction(Program.CurrentUser.Username,
                            "DeleteUser", $"删除用户：{row.Username}");
                        LoadData();
                        XtraMessageBox.Show("删除成功！", "提示");
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.LogError("删除用户失败", ex);
                    XtraMessageBox.Show("删除用户失败：" + ex.Message, "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text))
            {
                XtraMessageBox.Show("请输入用户名", "提示");
                return;
            }

            if (currentUserId == 0 && string.IsNullOrEmpty(txtPassword.Text))
            {
                XtraMessageBox.Show("请输入密码", "提示");
                return;
            }

            if (string.IsNullOrEmpty(cmbRole.Text))
            {
                XtraMessageBox.Show("请选择角色", "提示");
                return;
            }

            try
            {
                var user = new User
                {
                    Id = currentUserId,
                    Username = txtUsername.Text,
                    Role = cmbRole.Text,
                    IsActive = chkIsActive.Checked
                };

                if (!string.IsNullOrEmpty(txtPassword.Text))
                {
                    user.Password = EncryptPassword(txtPassword.Text);
                }

                var userDao = new UserDao();
                bool success;
                string action;

                if (currentUserId == 0)
                {
                    success = userDao.AddUser(user);
                    action = "AddUser";
                }
                else
                {
                    success = userDao.UpdateUser(user);
                    action = "UpdateUser";
                }

                if (success)
                {
                    LogHelper.LogUserAction(Program.CurrentUser.Username,
                        action, $"操作用户：{user.Username}");
                    LoadData();
                    SetEditMode(false);
                    XtraMessageBox.Show("保存成功！", "提示");
                }
            }
            catch (Exception ex)
            {
                LogHelper.LogError("保存用户失败", ex);
                XtraMessageBox.Show("保存失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            SetEditMode(false);
            ClearInputs();
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private string EncryptPassword(string password)
        {
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                var hashedBytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(password));
                return Convert.ToBase64String(hashedBytes);
            }
        }
    }
}