"""
主窗口界面
"""
import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QComboBox, QTextEdit, QMessageBox,
    QFileDialog, QGroupBox, QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from config_manager import ConfigManager
from database_manager import DatabaseManager
from excel_manager import ExcelManager
from email_manager import EmailManager
from template_config_dialog import TemplateConfigDialog


class BillGeneratorThread(QThread):
    """账单生成线程"""
    
    finished = pyqtSignal(bool, str)  # 完成信号: (成功/失败, 消息)
    progress = pyqtSignal(str)  # 进度信号
    
    def __init__(self, task_id, template_name, config_manager):
        super().__init__()
        self.task_id = task_id
        self.template_name = template_name
        self.config_manager = config_manager
    
    def run(self):
        """执行账单生成流程"""
        try:
            # 解析task_id，支持多个task_id用逗号分隔
            task_ids = [tid.strip() for tid in self.task_id.split(',') if tid.strip()]
            if not task_ids:
                self.finished.emit(False, "Task ID不能为空")
                return
            
            is_batch = len(task_ids) > 1
            if is_batch:
                self.progress.emit(f"检测到批量处理模式，共 {len(task_ids)} 个Task ID")
            
            # 1. 复制数据库到本地
            self.progress.emit("正在复制数据库到本地...")
            db_path = self.config_manager.get_database_path()
            local_db_path = self.config_manager.get_local_database_path()
            
            if not db_path:
                self.finished.emit(False, "未配置数据库路径，请在设置中配置")
                return
            
            db_manager = DatabaseManager(db_path, local_db_path)
            if not db_manager.copy_database_to_local():
                self.finished.emit(False, "复制数据库失败")
                return
            
            # 2. 连接数据库并查询数据
            self.progress.emit("正在查询数据库...")
            if not db_manager.connect():
                self.finished.emit(False, "连接数据库失败")
                return
            
            # 根据是否批量处理选择不同的数据准备方法
            if is_batch:
                task_data = db_manager.prepare_data_for_excel_batch(task_ids)
                staff_emails = db_manager.get_staff_emails_by_task_ids(task_ids)
            else:
                task_data = db_manager.prepare_data_for_excel(task_ids[0])
                staff_emails = db_manager.get_staff_emails_by_task_id(task_ids[0])
            
            if not task_data:
                db_manager.disconnect()
                self.finished.emit(False, f"未找到Task ID数据")
                return
            
            # 获取流水号
            serial_number = task_data.get('serial_number', '')
            
            # 生成Task ID显示文本（用于文件名和邮件主题）
            if is_batch:
                task_id_display = ', '.join(task_ids)
            else:
                task_id_display = task_ids[0]
            
            db_manager.disconnect()
            
            # 3. 生成Excel账单
            self.progress.emit("正在生成Excel账单...")
            template_config = self.config_manager.get_template(self.template_name)
            if not template_config:
                self.finished.emit(False, f"未找到模板配置: {self.template_name}")
                return
            
            excel_template_path = template_config.get("excel_template_path", "")
            if not excel_template_path or not os.path.exists(excel_template_path):
                self.finished.emit(False, f"Excel模板文件不存在: {excel_template_path}")
                return
            
            field_mappings = template_config.get("field_mappings", {})
            
            # 生成输出文件名
            output_folder = self.config_manager.get_output_folder()
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            # 文件名中的Task ID部分（批量时使用第一个和最后一个）
            if is_batch:
                filename_task_id = f"{task_ids[0]}~{task_ids[-1]}"
            else:
                filename_task_id = task_ids[0]
            
            # 文件名前缀改为"Audi China"
            output_filename = f"Audi China汽车维修服务询价：Task ID_{filename_task_id} - {serial_number}.xlsx"
            output_path = os.path.join(output_folder, output_filename)
            
            excel_manager = ExcelManager()
            if not excel_manager.create_bill_from_template(
                excel_template_path, output_path, task_data, field_mappings
            ):
                self.finished.emit(False, "生成Excel账单失败")
                return
            
            # 4. 创建Outlook邮件草稿
            self.progress.emit("正在创建Outlook邮件草稿...")
            email_manager = EmailManager()
            
            # 构建收件人列表
            to_recipients = []
            # 首位固定收件人组
            to_first = template_config.get("email_to_first", [])
            # 兼容旧版本配置（如果是字符串，转换为列表）
            if isinstance(to_first, str):
                to_first = [to_first] if to_first else []
            to_recipients.extend(to_first)
            # 动态收件人
            to_recipients.extend(staff_emails)
            # 末尾固定收件人组
            to_fixed = template_config.get("email_to_fixed", [])
            to_recipients.extend(to_fixed)
            
            # 抄送人列表
            cc_recipients = template_config.get("email_cc", [])
            
            # 邮件主题 - 前缀改为"Audi China"
            subject = f"Audi China汽车维修服务询价：Task ID:{task_id_display} - {serial_number}"
            
            # 邮件正文 - 使用模板生成
            task_description = task_data.get('task_description', '')
            bg_description = task_data.get('bg_description', '')
            
            # 获取邮件正文模板
            email_body_template = template_config.get("email_body_template", "")
            if not email_body_template:
                # 使用默认模板
                email_body_template = """hello
请参考附件服务 {task_description}

背景：{bg_description}

请2个工作日内确认是否可以提供服务，如果可以，请补充报价信息回复询价单。
预估到货时间，并附上报价依据截图。"""
            
            # 生成邮件正文
            body = EmailManager.generate_email_body_from_template(
                email_body_template,
                task_description,
                bg_description,
                task_ids if is_batch else None
            )
            
            # 附件列表
            attachments = [output_path]
            tutorial_path = self.config_manager.get_tutorial_attachment_path()
            if tutorial_path and os.path.exists(tutorial_path):
                attachments.append(tutorial_path)
            
            if not email_manager.create_draft_email(
                subject, to_recipients, cc_recipients, body, attachments
            ):
                self.finished.emit(False, "创建邮件草稿失败")
                return
            
            # 成功
            message = f"账单生成成功！\n\n文件路径: {output_path}\n\nOutlook邮件草稿已创建。"
            self.finished.emit(True, message)
            
        except Exception as e:
            self.finished.emit(False, f"处理过程中发生错误: {str(e)}")


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.generator_thread = None
        
        self.init_ui()
        self.load_templates()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("汽车维修服务账单生成器")
        self.setMinimumSize(800, 600)
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel("汽车维修服务账单生成器")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 基本设置组
        settings_group = QGroupBox("基本设置")
        settings_layout = QVBoxLayout()
        
        # 数据库路径
        db_layout = QHBoxLayout()
        db_layout.addWidget(QLabel("数据库路径:"))
        self.db_path_edit = QLineEdit()
        self.db_path_edit.setText(self.config_manager.get_database_path())
        db_layout.addWidget(self.db_path_edit)
        db_browse_btn = QPushButton("浏览...")
        db_browse_btn.clicked.connect(self.browse_database)
        db_layout.addWidget(db_browse_btn)
        settings_layout.addLayout(db_layout)
        
        # 输出文件夹
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件夹:"))
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setText(self.config_manager.get_output_folder())
        output_layout.addWidget(self.output_path_edit)
        output_browse_btn = QPushButton("浏览...")
        output_browse_btn.clicked.connect(self.browse_output_folder)
        output_layout.addWidget(output_browse_btn)
        settings_layout.addLayout(output_layout)
        
        # 教程附件路径
        tutorial_layout = QHBoxLayout()
        tutorial_layout.addWidget(QLabel("教程附件:"))
        self.tutorial_path_edit = QLineEdit()
        self.tutorial_path_edit.setText(self.config_manager.get_tutorial_attachment_path())
        tutorial_layout.addWidget(self.tutorial_path_edit)
        tutorial_browse_btn = QPushButton("浏览...")
        tutorial_browse_btn.clicked.connect(self.browse_tutorial)
        tutorial_layout.addWidget(tutorial_browse_btn)
        settings_layout.addLayout(tutorial_layout)
        
        # 保存设置按钮
        save_settings_btn = QPushButton("保存设置")
        save_settings_btn.clicked.connect(self.save_settings)
        settings_layout.addWidget(save_settings_btn)
        
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)
        
        # 模板管理组
        template_group = QGroupBox("模板管理")
        template_layout = QVBoxLayout()
        
        # 模板列表
        self.template_list = QListWidget()
        self.template_list.setMaximumHeight(150)
        template_layout.addWidget(self.template_list)
        
        # 模板管理按钮
        template_btn_layout = QHBoxLayout()
        add_template_btn = QPushButton("添加模板")
        add_template_btn.clicked.connect(self.add_template)
        template_btn_layout.addWidget(add_template_btn)
        
        edit_template_btn = QPushButton("编辑模板")
        edit_template_btn.clicked.connect(self.edit_template)
        template_btn_layout.addWidget(edit_template_btn)
        
        delete_template_btn = QPushButton("删除模板")
        delete_template_btn.clicked.connect(self.delete_template)
        template_btn_layout.addWidget(delete_template_btn)
        
        template_btn_layout.addStretch()
        template_layout.addLayout(template_btn_layout)
        
        template_group.setLayout(template_layout)
        main_layout.addWidget(template_group)
        
        # 生成账单组
        generate_group = QGroupBox("生成账单")
        generate_layout = QVBoxLayout()
        
        # Task ID输入
        task_id_layout = QHBoxLayout()
        task_id_layout.addWidget(QLabel("Task ID:"))
        self.task_id_edit = QLineEdit()
        self.task_id_edit.setPlaceholderText("请输入Task ID (支持多个，用逗号分隔，如: TASK001, TASK002)")
        task_id_layout.addWidget(self.task_id_edit)
        generate_layout.addLayout(task_id_layout)
        
        # 模板选择
        template_select_layout = QHBoxLayout()
        template_select_layout.addWidget(QLabel("选择模板:"))
        self.template_combo = QComboBox()
        template_select_layout.addWidget(self.template_combo)
        generate_layout.addLayout(template_select_layout)
        
        # 生成按钮
        self.generate_btn = QPushButton("生成询价单并创建邮件")
        self.generate_btn.setStyleSheet("font-size: 14px; padding: 10px; background-color: #4CAF50; color: white;")
        self.generate_btn.clicked.connect(self.generate_bill)
        generate_layout.addWidget(self.generate_btn)
        
        generate_group.setLayout(generate_layout)
        main_layout.addWidget(generate_group)
        
        # 日志输出
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        main_layout.addStretch()
        
        central_widget.setLayout(main_layout)
    
    def load_templates(self):
        """加载模板列表"""
        self.template_list.clear()
        self.template_combo.clear()
        
        templates = self.config_manager.get_templates()
        for template_name in templates.keys():
            self.template_list.addItem(template_name)
            self.template_combo.addItem(template_name)
    
    def browse_database(self):
        """浏览数据库文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择数据库文件",
            "",
            "数据库文件 (*.db *.sqlite)"
        )
        if file_path:
            self.db_path_edit.setText(file_path)
    
    def browse_output_folder(self):
        """浏览输出文件夹"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "选择输出文件夹"
        )
        if folder_path:
            self.output_path_edit.setText(folder_path)
    
    def browse_tutorial(self):
        """浏览教程附件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择教程附件",
            "",
            "所有文件 (*.*)"
        )
        if file_path:
            self.tutorial_path_edit.setText(file_path)
    
    def save_settings(self):
        """保存设置"""
        self.config_manager.set_database_path(self.db_path_edit.text())
        self.config_manager.set_output_folder(self.output_path_edit.text())
        self.config_manager.set_tutorial_attachment_path(self.tutorial_path_edit.text())
        
        QMessageBox.information(self, "成功", "设置已保存")
        self.log("设置已保存")
    
    def add_template(self):
        """添加模板"""
        dialog = TemplateConfigDialog(self)
        if dialog.exec():
            config = dialog.get_config()
            if config:
                template_name = config.pop("template_name")
                self.config_manager.add_template(template_name, config)
                self.load_templates()
                QMessageBox.information(self, "成功", f"模板 '{template_name}' 已添加")
                self.log(f"添加模板: {template_name}")
    
    def edit_template(self):
        """编辑模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择要编辑的模板")
            return
        
        template_name = current_item.text()
        template_config = self.config_manager.get_template(template_name)
        
        dialog = TemplateConfigDialog(self, template_name, template_config)
        if dialog.exec():
            config = dialog.get_config()
            if config:
                config.pop("template_name")
                self.config_manager.update_template(template_name, config)
                QMessageBox.information(self, "成功", f"模板 '{template_name}' 已更新")
                self.log(f"更新模板: {template_name}")
    
    def delete_template(self):
        """删除模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择要删除的模板")
            return
        
        template_name = current_item.text()
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除模板 '{template_name}' 吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.config_manager.delete_template(template_name)
            self.load_templates()
            QMessageBox.information(self, "成功", f"模板 '{template_name}' 已删除")
            self.log(f"删除模板: {template_name}")
    
    def generate_bill(self):
        """生成账单"""
        # 验证输入
        task_id_input = self.task_id_edit.text().strip()
        if not task_id_input:
            QMessageBox.warning(self, "错误", "请输入Task ID")
            return
        
        template_name = self.template_combo.currentText()
        if not template_name:
            QMessageBox.warning(self, "错误", "请选择模板")
            return
        
        # 解析task_id
        task_ids = [tid.strip() for tid in task_id_input.split(',') if tid.strip()]
        if len(task_ids) > 1:
            self.log(f"检测到批量模式: {len(task_ids)} 个Task ID")
        
        # 禁用生成按钮
        self.generate_btn.setEnabled(False)
        self.log_text.clear()
        self.log(f"开始生成询价单 - Task ID: {task_id_input}, 模板: {template_name}")
        
        # 创建并启动生成线程
        self.generator_thread = BillGeneratorThread(task_id_input, template_name, self.config_manager)
        self.generator_thread.progress.connect(self.log)
        self.generator_thread.finished.connect(self.on_generate_finished)
        self.generator_thread.start()
    
    def on_generate_finished(self, success: bool, message: str):
        """生成完成回调"""
        self.generate_btn.setEnabled(True)
        
        if success:
            self.log("✓ " + message)
            # 移除成功提示弹窗，只在日志中显示
        else:
            self.log("✗ " + message)
            QMessageBox.critical(self, "错误", message)
    
    def log(self, message: str):
        """添加日志"""
        self.log_text.append(message)
