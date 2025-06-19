using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using MoleLaboratoryExcel;
using System;
using System.Data;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

public partial class UserEditForm : XtraForm
{
    private int userId;
    private TextEdit txtUsername;
    private TextEdit txtPassword;
    private ComboBoxEdit cmbRole;
    private CheckEdit chkIsActive;
    private SimpleButton btnSave;
    private SimpleButton btnCancel;

    public UserEditForm(int id = 0)
    {
        userId = id;
        //InitializeComponent();
        InitializeUI();
        if (userId > 0)
        {
            LoadUserData();
        }
    }

    private void InitializeUI()
    {
        this.Text = userId > 0 ? "编辑用户" : "新增用户";
        this.Size = new System.Drawing.Size(300, 250);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // 创建控件
        var lblUsername = new LabelControl
        {
            Text = "用户名：",
            Location = new System.Drawing.Point(20, 20)
        };

        txtUsername = new TextEdit
        {
            Location = new System.Drawing.Point(100, 20),
            Size = new System.Drawing.Size(150, 20)
        };

        var lblPassword = new LabelControl
        {
            Text = "密码：",
            Location = new System.Drawing.Point(20, 50)
        };

        txtPassword = new TextEdit
        {
            Location = new System.Drawing.Point(100, 50),
            Size = new System.Drawing.Size(150, 20),
            Properties = { UseSystemPasswordChar = true }
        };

        var lblRole = new LabelControl
        {
            Text = "角色：",
            Location = new System.Drawing.Point(20, 80)
        };

        cmbRole = new ComboBoxEdit
        {
            Location = new System.Drawing.Point(100, 80),
            Size = new System.Drawing.Size(150, 20)
        };
        cmbRole.Properties.Items.AddRange(new string[] { "管理员", "普通用户" });

        chkIsActive = new CheckEdit
        {
            Text = "启用账号",
            Location = new System.Drawing.Point(100, 110),
            Checked = true
        };

        btnSave = new SimpleButton
        {
            Text = "保存",
            Location = new System.Drawing.Point(60, 150),
            DialogResult = DialogResult.OK
        };
        btnSave.Click += BtnSave_Click;

        btnCancel = new SimpleButton
        {
            Text = "取消",
            Location = new System.Drawing.Point(160, 150),
            DialogResult = DialogResult.Cancel
        };

        // 添加控件到窗体
        this.Controls.AddRange(new Control[]
        {
            lblUsername, txtUsername,
            lblPassword, txtPassword,
            lblRole, cmbRole,
            chkIsActive,
            btnSave, btnCancel
        });
    }

    private void LoadUserData()
    {
        // TODO: 从数据库加载用户数据
        // 这里使用示例数据
        txtUsername.Text = "admin";
        cmbRole.SelectedItem = "管理员";
        chkIsActive.Checked = true;
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(txtUsername.Text))
        {
            XtraMessageBox.Show("请输入用户名", "提示");
            return;
        }

        if (userId == 0 && string.IsNullOrEmpty(txtPassword.Text))
        {
            XtraMessageBox.Show("请输入密码", "提示");
            return;
        }

        if (cmbRole.SelectedItem == null)
        {
            XtraMessageBox.Show("请选择角色", "提示");
            return;
        }

        try
        {
            var user = new User
            {
                Id = userId,
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

            if (userId == 0)
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
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
        catch (Exception ex)
        {
            XtraMessageBox.Show("保存用户失败：" + ex.Message, "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // 密码加密方法
    private string EncryptPassword(string password)
    {
        using (var sha256 = SHA256.Create())
        {
            var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
            return Convert.ToBase64String(hashedBytes);
        }
    }
} 