using DevExpress.XtraEditors;
using System;
using System.Windows.Forms;

public class DatabaseConfigForm : XtraForm
{
    private TextEdit txtServer;
    private TextEdit txtDatabase;
    private TextEdit txtUsername;
    private TextEdit txtPassword;
    private SimpleButton btnTest;
    private SimpleButton btnSave;
    private SimpleButton btnCancel;

    public DatabaseConfigForm()
    {
        //InitializeComponent();
        InitializeUI();
    }

    private void InitializeUI()
    {
        this.Text = "数据库配置";
        this.Size = new System.Drawing.Size(400, 300);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // 创建控件
        var lblServer = new LabelControl { Text = "服务器：", Location = new System.Drawing.Point(30, 30) };
        txtServer = new TextEdit { Location = new System.Drawing.Point(100, 27), Width = 250 };

        var lblDatabase = new LabelControl { Text = "数据库：", Location = new System.Drawing.Point(30, 70) };
        txtDatabase = new TextEdit { Location = new System.Drawing.Point(100, 67), Width = 250 };

        var lblUsername = new LabelControl { Text = "用户名：", Location = new System.Drawing.Point(30, 110) };
        txtUsername = new TextEdit { Location = new System.Drawing.Point(100, 107), Width = 250 };

        var lblPassword = new LabelControl { Text = "密码：", Location = new System.Drawing.Point(30, 150) };
        txtPassword = new TextEdit 
        { 
            Location = new System.Drawing.Point(100, 147), 
            Width = 250,
            Properties = { PasswordChar = '*' }
        };

        btnTest = new SimpleButton
        {
            Text = "测试连接",
            Location = new System.Drawing.Point(100, 190),
            Width = 80
        };
        btnTest.Click += BtnTest_Click;

        btnSave = new SimpleButton
        {
            Text = "保存",
            Location = new System.Drawing.Point(190, 190),
            Width = 80
        };
        btnSave.Click += BtnSave_Click;

        btnCancel = new SimpleButton
        {
            Text = "取消",
            Location = new System.Drawing.Point(280, 190),
            Width = 80
        };
        btnCancel.Click += (s, e) => this.Close();

        // 添加控件
        this.Controls.AddRange(new Control[] {
            lblServer, txtServer,
            lblDatabase, txtDatabase,
            lblUsername, txtUsername,
            lblPassword, txtPassword,
            btnTest, btnSave, btnCancel
        });
    }

    private void BtnTest_Click(object sender, EventArgs e)
    {
        if (ValidateInputs())
        {
            string testConnectionString = BuildConnectionString();
            if (DbHelper.TestConnection(testConnectionString))
            {
                XtraMessageBox.Show("连接成功！", "提示", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                XtraMessageBox.Show("连接失败，请检查配置信息！", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        if (ValidateInputs())
        {
            try
            {
                DbHelper.UpdateConnectionString(
                    txtServer.Text.Trim(),
                    txtDatabase.Text.Trim(),
                    txtUsername.Text.Trim(),
                    txtPassword.Text
                );

                // 测试新的连接
                if (DbHelper.TestConnection())
                {
                    XtraMessageBox.Show("配置保存成功！", "提示");
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    XtraMessageBox.Show("配置保存失败，请检查连接信息！", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("保存配置失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private bool ValidateInputs()
    {
        if (string.IsNullOrWhiteSpace(txtServer.Text))
        {
            XtraMessageBox.Show("请输入服务器地址！", "提示");
            return false;
        }
        if (string.IsNullOrWhiteSpace(txtDatabase.Text))
        {
            XtraMessageBox.Show("请输入数据库名称！", "提示");
            return false;
        }
        if (string.IsNullOrWhiteSpace(txtUsername.Text))
        {
            XtraMessageBox.Show("请输入用户名！", "提示");
            return false;
        }
        if (string.IsNullOrWhiteSpace(txtPassword.Text))
        {
            XtraMessageBox.Show("请输入密码！", "提示");
            return false;
        }
        return true;
    }

    private string BuildConnectionString()
    {
        return $"Data Source={txtServer.Text.Trim()};Initial Catalog={txtDatabase.Text.Trim()};User ID={txtUsername.Text.Trim()};Password={txtPassword.Text};Connect Timeout=30;";
    }
} 