using DevExpress.XtraEditors;
using DevExpress.Utils;
using DevExpress.XtraLayout;
using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using MoleLaboratoryExcel;

namespace MoleLaboratoryExcel
{
    public partial class LoginForm : XtraForm
    {
        private TextEdit txtUsername;
        private TextEdit txtPassword;
        private SimpleButton btnLogin;
        private SimpleButton btnCancel;
        private CheckEdit chkRemember;
        private LabelControl lblTitle;
        private LabelControl lblSubTitle;
        private LayoutControl layoutControl;

        public LoginForm()
        {
            //InitializeComponent();
            InitializeUI();
            LoadSavedCredentials();
        }

        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new System.Drawing.Size(400, 300);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;

            // 创建布局控件
            layoutControl = new LayoutControl();
            layoutControl.Dock = DockStyle.Fill;

            // 创建标题
            lblTitle = new LabelControl
            {
                Text = "登录",
                Font = new Font("Microsoft YaHei", 20F, FontStyle.Bold),
                AutoSizeMode = LabelAutoSizeMode.None,
                Appearance = { TextOptions = { HAlignment = DevExpress.Utils.HorzAlignment.Center } }
            };

            //lblSubTitle = new LabelControl
            //{
            //    Text = "使用本地账号登录",
            //    Font = new Font("Microsoft YaHei", 10F),
            //    ForeColor = Color.Gray,
            //    AutoSizeMode = LabelAutoSizeMode.None,
            //    Appearance = { TextOptions = { HAlignment = DevExpress.Utils.HorzAlignment.Center } }
            //};

            // 创建输入控件
            txtUsername = new TextEdit();
            txtUsername.Text = "";
            txtUsername.Properties.NullText = "请输入用户名";
            txtUsername.Properties.Appearance.Font = new Font("Microsoft YaHei", 10F);
            txtUsername.Properties.Appearance.Options.UseFont = true;
            txtUsername.Properties.Appearance.Options.UseTextOptions = true;
            txtUsername.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;

            txtPassword = new TextEdit();
            txtPassword.Text = "";
            txtPassword.Properties.UseSystemPasswordChar = true;
            txtPassword.Properties.NullText = "请输入密码";
            txtPassword.Properties.Appearance.Font = new Font("Microsoft YaHei", 10F);
            txtPassword.Properties.Appearance.Options.UseFont = true;
            txtPassword.Properties.Appearance.Options.UseTextOptions = true;
            txtPassword.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;

            chkRemember = new CheckEdit
            {
                Text = "记住密码",
                Font = new Font("Microsoft YaHei", 9F)
              
            };
            chkRemember.Properties.AllowGrayed = false;
            chkRemember.Properties.Caption = "记住密码";

            btnLogin = new SimpleButton();
            btnLogin.Text = "登录";
            btnLogin.Font = new Font("Microsoft YaHei", 10F);
            btnLogin.Appearance.BackColor = Color.FromArgb(103, 58, 183); // 紫色
            btnLogin.Appearance.ForeColor = Color.White;
            btnLogin.Appearance.Options.UseBackColor = true;
            btnLogin.Appearance.Options.UseForeColor = true;
            btnLogin.Size = new Size(280, 40);
            btnLogin.Click += (s, e) => HandleLogin();

            btnCancel = new SimpleButton();
            btnCancel.Text = "取消";
            btnCancel.Appearance.Font = new Font("Microsoft YaHei", 10F);
            btnCancel.Size = new Size(100, 35);
            btnCancel.Click += BtnCancel_Click;

            // 添加到布局
            var itemTitle = layoutControl.AddItem(string.Empty, lblTitle);
            itemTitle.TextVisible = false;
            itemTitle.SizeConstraintsType = SizeConstraintsType.Custom;
            itemTitle.MinSize = new Size(0, 40);
            itemTitle.MaxSize = new Size(0, 40);

            var itemSubTitle = layoutControl.AddItem(string.Empty, lblSubTitle);
            itemSubTitle.TextVisible = false;
            itemSubTitle.SizeConstraintsType = SizeConstraintsType.Custom;
            itemSubTitle.MinSize = new Size(0, 30);
            itemSubTitle.MaxSize = new Size(0, 30);

            var itemUsername = layoutControl.AddItem(string.Empty, txtUsername);
            itemUsername.TextVisible = false;
            itemUsername.Padding = new DevExpress.XtraLayout.Utils.Padding(50, 10, 5, 5);

            var itemPassword = layoutControl.AddItem(string.Empty, txtPassword);
            itemPassword.TextVisible = false;
            itemPassword.Padding = new DevExpress.XtraLayout.Utils.Padding(50, 10, 5, 5);

            var itemRemember = layoutControl.AddItem(string.Empty, chkRemember);
            itemRemember.TextVisible = false;          
            itemRemember.Padding = new DevExpress.XtraLayout.Utils.Padding(50, 10, 5, 15);

            var itemLogin = layoutControl.AddItem(string.Empty, btnLogin);
            itemLogin.TextVisible = false;
            itemLogin.Padding = new DevExpress.XtraLayout.Utils.Padding(50, 10, 5, 5);

            // 设置布局组的属性
            layoutControl.Root.GroupBordersVisible = false;
            layoutControl.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            layoutControl.Root.TextVisible = false;

            // 设置回车键事件
            txtPassword.KeyPress += (s, e) =>
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    e.Handled = true;
                    HandleLogin();
                }
            };

            txtUsername.KeyPress += (s, e) =>
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    e.Handled = true;
                    txtPassword.Focus();
                }
            };

            // 添加布局到窗体
            this.Controls.Add(layoutControl);
            this.ActiveControl = txtUsername;
        }

        private void LoadSavedCredentials()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\MoleLaboratoryExcel"))
                {
                    if (key != null)
                    {
                        var remember = key.GetValue("RememberPassword")?.ToString() == "1";
                        if (remember)
                        {
                            var username = key.GetValue("SavedUsername")?.ToString();
                            var encryptedPassword = key.GetValue("SavedPassword")?.ToString();

                            if (!string.IsNullOrEmpty(username))
                            {
                                txtUsername.Text = username;
                                if (!string.IsNullOrEmpty(encryptedPassword))
                                {
                                    txtPassword.Text = DecryptPassword(encryptedPassword);
                                }
                                chkRemember.Checked = true;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void SaveCredentials()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\MoleLaboratoryExcel"))
                {
                    if (chkRemember.Checked)
                    {
                        key.SetValue("SavedUsername", txtUsername.Text);
                        key.SetValue("SavedPassword", EncryptPassword(txtPassword.Text));
                        key.SetValue("RememberPassword", "1");
                    }
                    else
                    {
                        key.DeleteValue("SavedUsername", false);
                        key.DeleteValue("SavedPassword", false);
                        key.SetValue("RememberPassword", "0");
                    }
                }
            }
            catch { }
        }

        private string EncryptPassword(string password)
        {
            try
            {
                byte[] data = System.Text.Encoding.UTF8.GetBytes(password);
                // 使用更安全的加密方式
                using (var aes = System.Security.Cryptography.Aes.Create())
                {
                    aes.Key = new byte[32] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32 };
                    aes.IV = new byte[16] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };

                    using (var encryptor = aes.CreateEncryptor())
                    {
                        byte[] encryptedData = encryptor.TransformFinalBlock(data, 0, data.Length);
                        return Convert.ToBase64String(encryptedData);
                    }
                }
            }
            catch
            {
                return "";
            }
        }

        private string DecryptPassword(string encryptedPassword)
        {
            try
            {
                byte[] data = Convert.FromBase64String(encryptedPassword);
                using (var aes = System.Security.Cryptography.Aes.Create())
                {
                    aes.Key = new byte[32] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32 };
                    aes.IV = new byte[16] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };

                    using (var decryptor = aes.CreateDecryptor())
                    {
                        byte[] decryptedData = decryptor.TransformFinalBlock(data, 0, data.Length);
                        return System.Text.Encoding.UTF8.GetString(decryptedData);
                    }
                }
            }
            catch
            {
                return "";
            }
        }

        private void HandleLogin()
        {
            if (string.IsNullOrEmpty(txtUsername.Text))
            {
                XtraMessageBox.Show("请输入用户名", "提示");
                txtUsername.Focus();
                return;
            }

            if (string.IsNullOrEmpty(txtPassword.Text))
            {
                XtraMessageBox.Show("请输入密码", "提示");
                txtPassword.Focus();
                return;
            }

            try
            {
                var userDao = new UserDao();
                var user = userDao.GetUserByUsername(txtUsername.Text);

                if (user != null && ValidatePassword(txtPassword.Text, user.Password))
                {
                    SaveCredentials(); // 保存凭据
                    Program.CurrentUser = user;
                    LogHelper.LogUserAction(user.Username, "Login", "用户登录成功");
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    LogHelper.LogUserAction(txtUsername.Text, "Error", "登录失败：用户名或密码错误");
                    XtraMessageBox.Show("用户名或密码错误", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtPassword.SelectAll();
                    txtPassword.Focus();
                }
            }
            catch (Exception ex)
            {
                LogHelper.LogError("登录失败", ex);
                XtraMessageBox.Show("登录失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private bool ValidatePassword(string inputPassword, string storedPassword)
        {
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                var hashedBytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(inputPassword));
                string hashedPassword = Convert.ToBase64String(hashedBytes);
                return hashedPassword == storedPassword;
            }
        }
    }
}