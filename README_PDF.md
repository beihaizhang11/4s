# 汽车维修服务询价邮件生成器 (PDF变体版本)

这是PDF变体版本，专门用于扫描PDF文件并通过邮件发送，不使用Excel模板。

## 与标准版本的区别

| 功能 | 标准版本 | PDF版本 |
|------|---------|---------|
| **Excel模板** | ✅ 需要 | ❌ 不需要 |
| **生成Excel账单** | ✅ 生成 | ❌ 不生成 |
| **字段映射配置** | ✅ 需要配置 | ❌ 不需要 |
| **PDF扫描** | ❌ 无 | ✅ 扫描指定目录 |
| **邮件附件** | Excel账单 + 教程 | 所有PDF + 教程 |
| **邮件主题** | "Audi China汽车维修服务询价..." | "**已收货** - Audi China汽车维修服务询价..." |
| **配置目录** | 输出文件夹 | PDF文件目录 |

## 运行方式

```bash
# PDF版本
python main_pdf.py

# 标准版本
python main.py
```

## 邮件主题格式

```
已收货 - Audi China 汽车维修服务询价：Task ID: {task_id}
```

## 使用流程

1. **配置设置**
   - 数据库路径
   - PDF文件目录（指定包含PDF文件的文件夹）
   - 教程附件（可选）

2. **配置模板**
   - 添加模板
   - 配置邮件收件人和抄送人
   - 编辑邮件正文模板

3. **生成邮件**
   - 输入Task ID（支持多个，用逗号分隔）
   - 选择模板
   - 点击"生成询价邮件 (附PDF)"
   - Outlook自动打开邮件草稿

## PDF扫描规则

- 扫描指定目录下的所有`.pdf`文件（不区分大小写）
- 不扫描子目录
- 按文件名排序
- 所有找到的PDF文件都会作为邮件附件

## 附件内容

1. PDF文件目录中的所有PDF文件（自动扫描）
2. 教程附件（可选配置）

## 邮件收件人构成

**TO列表**:
1. 首位固定收件人组（模板配置）
2. sender_mail（动态，来自tasks表）
3. staff_email列表（动态，从tasks_staff表查询）
4. 末尾固定收件人组（模板配置）

**CC列表**: 全部固定，可在模板配置中修改

## 批量处理

支持一次输入多个Task ID（用逗号分隔）：
```
TASK001, TASK002, TASK003
```

所有PDF文件都会附加到同一封邮件中。

## 版本切换

```bash
# 切换到标准版本
git checkout cursor/automated-billing-and-email-c982

# 切换到PDF版本
git checkout pdf-attachment-variant

# 或直接运行不同的主程序（两个版本可共存）
python main.py      # 标准版本
python main_pdf.py  # PDF版本
```

## 系统要求

- Windows 10+
- Python 3.8+
- Microsoft Outlook
- 数据库访问权限

---

**PDF版本专用 - 收货后发送PDF文件！** 📄
