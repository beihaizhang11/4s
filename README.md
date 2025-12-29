# 汽车维修服务账单生成器

这是一个基于Python和PyQt6开发的桌面应用程序，用于自动化生成汽车维修服务账单并通过Outlook发送邮件。

## 功能特点

### 核心功能
- 📊 **自动生成账单**: 根据Task ID从数据库查询数据，自动填充Excel账单模板
- 📧 **邮件集成**: 自动创建Outlook邮件草稿，附带生成的账单和教程文件
- 🔢 **流水号生成**: 自动生成格式化的流水号 `WST_{car_model}_{carid}_{sender_name}_{sent_time}`
- 🔄 **数据库同步**: 从共享文件夹复制数据库到本地，避免影响在线数据库

### 模板管理
- ➕ **添加模板**: 支持添加多个账单模板
- ✏️ **编辑模板**: 配置数据库字段到Excel单元格的映射关系
- 🗺️ **灵活映射**: 支持一个字段映射到多个单元格
- 📮 **邮件配置**: 为每个模板配置独立的邮件收件人和抄送人

### 邮件配置
- **TO列表构成**:
  1. 首位固定收件人组（模板配置，可多个）
  2. 动态收件人（从tasks_staff表查询）
  3. 末尾固定收件人组（模板配置，可多个）
- **CC列表**: 全部固定，可在模板配置中修改

## 系统要求

- Windows 10 或更高版本
- Python 3.8 或更高版本
- Microsoft Outlook（已安装并配置）
- 数据库访问权限

## 安装步骤

1. **克隆或下载项目**
   ```bash
   git clone <repository_url>
   cd <project_directory>
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

3. **运行程序**
   ```bash
   python main.py
   ```

## 使用指南

### 首次使用配置

1. **配置基本设置**
   - **数据库路径**: 选择共享文件夹中的数据库文件路径
   - **输出文件夹**: 选择生成的账单保存位置
   - **教程附件**: 选择要随邮件一起发送的教程文件
   - 点击"保存设置"按钮

2. **添加账单模板**
   - 点击"添加模板"按钮
   - 输入模板名称（如："AUDI", "VW", "VGIC"）
   - 选择Excel模板文件
   - 配置字段映射（见下方详细说明）
   - 配置邮件收件人和抄送人
   - 点击"保存"

### 字段映射配置

字段映射定义了数据库字段如何写入Excel单元格。

**可用的数据库字段**:
- `task_id`: 任务ID
- `sender_name`: 发件人姓名
- `sender_mail`: 发件人邮箱
- `sender_phone`: 发件人电话
- `VIN`: 车辆识别号
- `task_dept`: 任务部门
- `sent_time`: 发送时间
- `task_description`: 任务描述
- `bg_description`: 背景描述
- `serial_number`: 自动生成的流水号

**单元格格式**:
- 简单引用: `A1`, `B2`, `C10`
- 指定工作表: `Sheet1!A1`, `数据!B2`
- 多个单元格: `A1, B2, Sheet1!C3` (用逗号分隔)

**示例**:
```
字段名: sender_name
单元格: A5, Sheet2!B10
说明: sender_name会同时写入当前工作表的A5单元格和Sheet2工作表的B10单元格
```

### 邮件配置说明

**TO首位收件人列表**:
- 输入多个邮箱地址，每行一个
- 这些收件人总是出现在TO列表的最前面

**TO固定收件人列表**:
- 输入多个邮箱地址，每行一个
- 这些收件人会添加在动态收件人之后

**CC抄送人列表**:
- 输入多个邮箱地址，每行一个
- 所有CC收件人都是固定的

**邮件收件人顺序**:
```
TO: [首位固定组(可多个)] → [动态查询的staff_email] → [末尾固定组(可多个)]
CC: [全部固定]
```

**邮件正文模板**:
邮件正文支持自定义模板，可在模板配置中编辑。

支持的变量：
- `{task_description}`: 任务描述
- `{bg_description}`: 背景描述（如果为空，包含此变量的行会自动删除）
- `{task_ids}`: 多个Task ID时显示，格式如"TASK001, TASK002"（单个Task ID时自动删除此行）

默认模板：
```
hello
请参考附件服务 {task_description}

背景：{bg_description}

请2个工作日内确认是否可以提供服务，如果可以，请补充报价信息回复询价单。
预估到货时间，并附上报价依据截图。
```

### 生成账单流程

1. **输入Task ID**: 在"Task ID"输入框中输入要查询的任务ID
   - 单个Task ID: 直接输入，如 `TASK001`
   - 多个Task ID: 用逗号分隔，如 `TASK001, TASK002, TASK003`
2. **选择模板**: 从下拉列表中选择要使用的账单模板
3. **生成账单**: 点击"生成账单并创建邮件"按钮
4. **查看结果**: 
   - 处理日志会显示每个步骤的执行情况
   - 成功后会显示生成的账单文件路径
   - Outlook会自动打开邮件草稿窗口
5. **检查邮件**: 在Outlook中检查邮件内容和附件，确认无误后发送

## 文件结构

```
project/
├── main.py                      # 主程序入口
├── main_window.py               # 主窗口界面
├── config_manager.py            # 配置管理模块
├── database_manager.py          # 数据库操作模块
├── excel_manager.py             # Excel操作模块
├── email_manager.py             # 邮件操作模块
├── template_config_dialog.py   # 模板配置对话框
├── requirements.txt             # 依赖包列表
├── config.json                  # 配置文件（自动生成）
├── output/                      # 输出文件夹（自动创建）
└── README.md                    # 本文档
```

## 数据库结构

程序使用SQLite数据库，主要涉及以下表：

### tasks 表
存储任务主信息，包含以下字段：
- `task_id`: 任务ID（主键）
- `sender_name`: 发件人姓名
- `sender_mail`: 发件人邮箱
- `sender_phone`: 发件人电话
- `VIN`: 车辆识别号
- `task_dept`: 任务部门
- `sent_time`: 发送时间
- `task_description`: 任务描述
- `bg_description`: 背景描述
- `car_model`: 车型
- `carid`: 车辆ID

### tasks_staff 表
存储任务关联的员工信息：
- `task_id`: 任务ID
- `staff_email`: 员工邮箱
- `user_account`: 用户账号
- `first_name`: 名字
- `last_name`: 姓氏

## 流水号规则

流水号格式: `WST_{car_model}_{carid}_{sender_name}_{sent_time}`

**示例**:
- 输入数据: 
  - car_model: "Audi A4"
  - carid: "12345"
  - sender_name: "张三"
  - sent_time: "2025-12-29 14:30:00"
- 生成流水号: `WST_Audi_A4_12345_张三_2025-12-29_14-30-00`

## 生成的文件命名

**账单文件名格式**:
```
Audi China汽车维修服务询价：Task ID_{task_id} - {流水号}.xlsx
```

**示例**:
```
Audi China汽车维修服务询价：Task ID_TASK001 - WST_Audi_A4_12345_张三_2025-12-29_14-30-00.xlsx
```

## 邮件主题格式

```
Audi China汽车维修服务询价：Task ID:{task_id} - {流水号}
```

**示例**:
```
Audi China汽车维修服务询价：Task ID:TASK001 - WST_Audi_A4_12345_张三_2025-12-29_14-30-00
```

## 注意事项

1. **Outlook要求**: 
   - 必须安装Microsoft Outlook
   - Outlook必须配置了邮箱账户
   - 首次运行可能需要授予权限

2. **数据库访问**:
   - 确保有权访问共享文件夹中的数据库
   - 程序会复制数据库到本地，不会修改源数据库

3. **Excel模板**:
   - 模板文件必须是.xlsx或.xls格式
   - 确保模板文件路径正确且可访问
   - 单元格引用区分大小写（如A1, Sheet1!B2）

4. **权限问题**:
   - 确保有权限读取数据库文件
   - 确保有权限写入输出文件夹
   - 确保有权限访问Excel模板文件

5. **配置文件**:
   - 配置保存在config.json文件中
   - 请勿手动编辑config.json，使用程序界面进行配置

## 常见问题

### Q: Outlook没有弹出邮件窗口？
A: 
1. 检查Outlook是否正在运行
2. 检查是否有防火墙或安全软件阻止
3. 尝试手动打开Outlook后再运行程序

### Q: 找不到Task ID对应的数据？
A: 
1. 确认Task ID输入正确
2. 确认数据库路径配置正确
3. 检查数据库中是否存在该Task ID的记录

### Q: Excel账单生成失败？
A: 
1. 检查Excel模板文件是否存在
2. 检查字段映射配置是否正确
3. 检查单元格引用格式是否正确（如A1, Sheet1!B2）

### Q: 邮件收件人列表不正确？
A: 
1. 检查模板配置中的邮件设置
2. 确认tasks_staff表中有该task_id的记录
3. 检查staff_email字段是否有效

## 技术支持

如有问题或建议，请联系技术支持团队。

## 批量处理功能

### 功能说明
支持一次输入多个Task ID（用逗号分隔），生成一个合并的Excel账单。

### 使用方法
在Task ID输入框中输入多个Task ID，用逗号分隔：
```
TASK001, TASK002, TASK003
```

### 数据合并规则
1. **task_id字段**: 所有Task ID用换行符连接，写入同一单元格
   ```
   TASK001
   TASK002
   TASK003
   ```

2. **VIN字段**: 所有VIN号用换行符连接，写入同一单元格
   ```
   VIN123456789
   VIN987654321
   VIN111222333
   ```

3. **serial_number字段**: 所有流水号用换行符连接，写入同一单元格
   ```
   WST_Audi_A4_12345_张三_2025-12-29_14-30-00
   WST_Audi_A6_67890_李四_2025-12-29_15-00-00
   WST_BMW_X5_11111_王五_2025-12-29_16-00-00
   ```

4. **其他字段**: 使用第一个Task ID的数据
   - sender_name, sender_mail, sender_phone
   - task_dept, sent_time
   - task_description, bg_description
   - 流水号（基于第一个Task ID生成）

4. **动态收件人**: 合并所有Task ID关联的staff_email（自动去重）

### 文件命名
文件名前缀统一为"Audi China"，批量处理时使用首尾Task ID：
```
Audi China汽车维修服务询价：Task ID_TASK001~TASK003 - {流水号}.xlsx
```

### 邮件主题
邮件主题前缀统一为"Audi China"：
```
Audi China汽车维修服务询价：Task ID:TASK001, TASK002, TASK003 - {流水号}
```

### 邮件正文
批量处理时会自动在邮件正文中标注包含的所有Task ID。

## 更新日志

### v1.2.0 (2025-12-29)
- 新增邮件正文模板功能，支持自定义邮件内容
- 文件名和邮件主题前缀统一为"Audi China"
- 批量处理时流水号也支持换行叠加
- 优化用户体验，移除成功提示弹窗
- 按钮文字更新为"生成询价单并创建邮件"

### v1.1.0 (2025-12-29)
- 新增批量处理功能，支持多个Task ID
- task_id和VIN字段支持多行数据写入
- 优化邮件正文格式
- 支持TO首位收件人添加多个

### v1.0.0 (2025-12-29)
- 初始版本发布
- 支持多模板管理
- 支持灵活的字段映射
- 自动生成流水号
- Outlook邮件集成
- 数据库本地化操作

## 许可证

本软件仅供内部使用。
