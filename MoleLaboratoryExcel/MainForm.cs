using DevExpress.XtraBars;
using DevExpress.XtraBars.FluentDesignSystem;
using DevExpress.XtraBars.Navigation;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Svg;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using MoleLaboratoryExcel.Forms;

namespace MoleLaboratoryExcel
{
    public partial class MainForm : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
    {
        public MainForm()
        {
            InitializeComponent();
            InitializeUI();
        }

        private void InitializeUI()
        {
            // 设置窗体属性
            this.Text = "默乐生物工艺部实验报告软件";
            this.WindowState = FormWindowState.Maximized;

            // 创建Excel导入组
            var excelGroup = new AccordionControlElement
            {
                Text = "Excel导入",
                Style = ElementStyle.Group
            };

            // 创建Word模板组
            var wordGroup = new AccordionControlElement
            {
                Text = "生成Word模板",
                Style = ElementStyle.Group
            };

            // 创建系统管理组
            var systemGroup = new AccordionControlElement
            {
                Text = "系统管理",
                Style = ElementStyle.Group
            };

            // 添加子项
            var excelItem1 = new AccordionControlElement
            {
                Style = ElementStyle.Item,
                Text = "导入Excel数据"
            };
            excelItem1.Click += (s, e) => HandleExcelImport();

            var wordItem1 = new AccordionControlElement
            {
                Style = ElementStyle.Item,
                Text = "生成报告模板"
            };
            wordItem1.Click += (s, e) => HandleWordTemplate();

            // 添加用户管理菜单项
            var userManageItem = new AccordionControlElement
            {
                Style = ElementStyle.Item,
                Text = "用户管理"
            };
            userManageItem.Click += (s, e) => HandleUserManage();

            // 添加日志查询菜单项
            var logQueryItem = new AccordionControlElement
            {
                Style = ElementStyle.Item,
                Text = "日志查询"
            };
            logQueryItem.Click += (s, e) => HandleLogQuery();

            // 组装菜单结构
            excelGroup.Elements.Add(excelItem1);
            wordGroup.Elements.Add(wordItem1);
            systemGroup.Elements.AddRange(new[] { userManageItem, logQueryItem });

            // 根据用户角色显示菜单
            var menuItems = new List<AccordionControlElement> { excelGroup, wordGroup };

            // 只有管理员可以看到系统管理菜单
            if (Program.CurrentUser?.Role == "管理员")
            {
                menuItems.Add(systemGroup);
            }

            // 添加菜单项到 AccordionControl
            accordionControl.Elements.AddRange(menuItems.ToArray());

            // 设置主题和颜色
            UserLookAndFeel.Default.SetSkinStyle(
                DevExpress.LookAndFeel.SkinStyle.Bezier,
                DevExpress.LookAndFeel.SkinSvgPalette.Bezier.Default);

            // 设置标题
            var captionItem = new BarStaticItem();
            captionItem.Caption = "默乐生物工艺部实验报告软件";
            fluentDesignFormControl.Items.Add(captionItem);

            // 添加用户信息显示
            var userItem = new BarStaticItem();
            userItem.Caption = $"当前用户：{Program.CurrentUser?.Username}";
            fluentDesignFormControl.Items.Add(userItem);

            // 添加退出按钮
            var logoutItem = new BarButtonItem(barManager, "退出");
            // 从嵌入资源加载 SVG 图标
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("MoleLaboratoryExcel.Images.actions_exit.svg"))
            {
                if (stream != null)
                {
                    logoutItem.ImageOptions.SvgImage = DevExpress.Utils.Svg.SvgImage.FromStream(stream);
                }
            }
            logoutItem.ItemClick += (s, e) => Application.Exit();
            fluentDesignFormControl.Items.Add(logoutItem);

            // 展开所有组
            accordionControl.ExpandAll();
        }

        private void HandleExcelImport()
        {
            var excelImportForm = new Forms.ExcelImportForm();
            excelImportForm.ShowDialog();
        }

        private void HandleWordTemplate()
        {
            var excelToWordForm = new Forms.ExcelToWordForm();
            excelToWordForm.ShowDialog();
        }

        private void HandleUserManage()
        {
            var userManageForm = new UserManageForm();
            userManageForm.StartPosition = FormStartPosition.CenterScreen;
            userManageForm.ShowDialog();
        }

        private void HandleLogQuery()
        {
            var logQueryForm = new LogQueryForm();
            logQueryForm.StartPosition = FormStartPosition.CenterScreen;
            logQueryForm.ShowDialog();
        }
    }
}