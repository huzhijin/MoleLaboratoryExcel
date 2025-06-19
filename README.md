# MoleLaboratoryExcel

## 项目简介

**MoleLaboratoryExcel** 是一个专为**默乐生物工艺部**开发的实验室数据管理软件，主要用于处理和整理实验仪器产生的Excel数据，并自动生成标准化的Word实验报告。

## 功能特色

### 🔬 多仪器数据支持
- **赛默飞7500 (ThermoFisher7500)**: 支持赛默飞仪器数据格式
- **宏石 (HONGSHI)**: 支持宏石仪器数据格式
- 智能识别不同仪器的数据起始行和格式

### 📊 Excel数据处理
- **批量文件处理**: 同时处理多个Excel文件
- **智能数据整理**: 
  - 分开整理：按工作表分别生成文件
  - 合并整理：将多个数据源合并处理
- **数据验证**: 自动检测和处理数据格式

### 📝 Word报告生成
- **标准化模板**: 自动生成符合实验室规范的Word报告
- **完整章节结构**:
  - 目的
  - 实验地点和时间
  - 实验方案
  - 结果与分析（自动插入Excel表格）
  - 结论
  - 参考文献
  - 附件
- **自动目录**: 生成带目录的专业报告
- **表格转换**: Excel数据无缝转换为Word表格

### 👥 用户管理系统
- **多用户支持**: 用户注册、登录、权限管理
- **角色分离**: 普通用户和管理员权限区分
- **操作审计**: 详细记录所有用户操作日志

## 技术栈

### 开发环境
- **.NET Framework 4.8**
- **WinForms** 桌面应用程序
- **Visual Studio 2017+**

### 核心依赖
- **DevExpress v24.1**: 现代化UI控件库
- **EPPlus v7.5.1**: Excel文件处理
- **NPOI v2.7.2**: Office文档操作
- **DocumentFormat.OpenXml v3.2.0**: Word文档生成
- **SQL Server**: 数据存储

### 主要NuGet包
```xml
<package id="EPPlus" version="7.5.1" />
<package id="NPOI" version="2.7.2" />
<package id="DocumentFormat.OpenXml" version="3.2.0" />
<package id="DevExpress.Win" version="24.1.6" />
```

## 系统要求

- **操作系统**: Windows 7/8/10/11
- **.NET Framework**: 4.8或更高版本
- **数据库**: SQL Server 2012或更高版本
- **DevExpress**: 需要有效的DevExpress许可证

## 安装和配置

### 1. 克隆项目
```bash
git clone https://github.com/huzhijin/MoleLaboratoryExcel.git
cd MoleLaboratoryExcel
```

### 2. 数据库配置
1. 在SQL Server中创建数据库
2. 执行 `MoleLaboratoryExcel/Database/CreateTables.sql` 创建必要的表
3. 修改 `App.config` 中的连接字符串

### 3. 编译运行
1. 使用Visual Studio打开 `MoleLaboratoryExcel.sln`
2. 还原NuGet包
3. 编译并运行项目

## 使用说明

### 首次使用
1. 启动应用程序
2. 配置数据库连接
3. 创建管理员账户
4. 开始使用

### 操作流程
1. **用户登录**: 输入用户名和密码
2. **选择仪器类型**: 根据数据来源选择对应仪器
3. **导入Excel文件**: 选择需要处理的Excel文件
4. **数据整理**: 选择分开整理或合并整理
5. **生成Word报告**: 自动生成标准化实验报告

## 项目结构

```
MoleLaboratoryExcel/
├── Data/           # 数据访问层
├── Models/         # 数据模型
├── Forms/          # 用户界面窗体
├── Utils/          # 工具类
├── Helpers/        # 辅助类
├── Database/       # 数据库脚本
└── Properties/     # 项目属性
```

## 开发团队

- **开发者**: [@huzhijin](https://github.com/huzhijin)
- **项目性质**: 企业内部使用软件

## 许可证

本项目为企业内部使用软件，版权所有。

## 联系方式

如有问题或建议，请联系开发团队。

---

**默乐生物工艺部实验报告软件** - 让实验数据管理更简单、更专业！ 