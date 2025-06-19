using DevExpress.Skins;
using DevExpress.UserSkins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace MoleLaboratoryExcel
{
    static class Program
    {
        public static User CurrentUser { get; set; }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 确保存在管理员账号
            EnsureAdminUser();

            var loginForm = new LoginForm();
            if (loginForm.ShowDialog() == DialogResult.OK)
            {
                Application.Run(new MainForm());
            }
        }

        private static void EnsureAdminUser()
        {
            try
            {
                var userDao = new UserDao();
                var adminUser = userDao.GetUserByUsername("admin");
                if (adminUser == null)
                {
                    // 创建默认管理员账号
                    var user = new User
                    {
                        Username = "admin",
                        Password = GetHashedPassword("123456"),
                        Role = "管理员",
                        IsActive = true,
                        CreateTime = DateTime.Now
                    };
                    userDao.AddUser(user);
                }
            }
            catch (Exception ex)
            {
                LogHelper.LogError("创建管理员账号失败", ex);
            }
        }

        private static string GetHashedPassword(string password)
        {
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                var hashedBytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(password));
                return Convert.ToBase64String(hashedBytes);
            }
        }
    }
}